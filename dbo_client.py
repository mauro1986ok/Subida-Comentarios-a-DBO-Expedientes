# -*- coding: utf-8 -*-
import os
import time
import logging
from typing import Optional, Dict, Any, Union, List
from dataclasses import dataclass
import requests
from dotenv import load_dotenv

load_dotenv()

logger = logging.getLogger(__name__)


class DBOException(Exception):
    """Excepción base para errores de la API de DBO."""
    pass


class DBOAuthenticationError(DBOException):
    """Error al autenticar contra la API de DBO."""
    pass


class DBONetworkError(DBOException):
    """Error de conexión o red con la API de DBO."""
    pass


class DBOAPIError(DBOException):
    """Error devuelto por la API de DBO (códigos 4xx, 5xx)."""
    def __init__(self, message: str, status_code: Optional[int] = None, response_text: Optional[str] = None):
        super().__init__(message)
        self.status_code = status_code
        self.response_text = response_text


@dataclass
class DBOConfig:
    username: str
    password: str
    base_url: str = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com"
    auth_endpoint: str = "/api/login"
    
    @classmethod
    def from_env(cls) -> "DBOConfig":
        username = os.getenv("DBO_USERNAME") or os.getenv("DBO_USUARIO")
        password = os.getenv("DBO_PASSWORD") or os.getenv("DBO_CONTRASENA")
        base_url = os.getenv("DBO_BASE_URL", "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com").rstrip("/")
        auth_endpoint = os.getenv("DBO_AUTH_ENDPOINT", "/api/login")
        
        if not username or not password:
            raise ValueError("Las variables DBO_USERNAME (o DBO_USUARIO) y DBO_PASSWORD (o DBO_CONTRASENA) deben estar definidas en el archivo .env o en el entorno.")
            
        return cls(
            username=username,
            password=password,
            base_url=base_url,
            auth_endpoint=auth_endpoint
        )


class DBOClient:
    def __init__(
        self,
        username: Optional[str] = None,
        password: Optional[str] = None,
        base_url: str = "http://dbo2dhartec-env.eba-as23ttdp.us-west-2.elasticbeanstalk.com",
        auth_endpoint: str = "/api/login",
        timeout: int = 30,
        max_retries: int = 3,
        retry_delay: float = 2.0,
        auto_login: bool = True
    ):
        self.username = username
        self.password = password
        self.base_url = base_url.rstrip("/")
        self.auth_endpoint = auth_endpoint
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay

        self.session = requests.Session()
        self._is_authenticated = False
        self._current_user: Optional[Dict[str, Any]] = None

        if auto_login and self.username and self.password:
            self.login()

    @classmethod
    def from_env(cls, **kwargs) -> "DBOClient":
        """Crea una instancia de DBOClient utilizando la configuración del entorno (.env)."""
        config = DBOConfig.from_env()
        return cls(
            username=config.username,
            password=config.password,
            base_url=kwargs.get("base_url", config.base_url),
            auth_endpoint=kwargs.get("auth_endpoint", config.auth_endpoint),
            **{k: v for k, v in kwargs.items() if k not in ("base_url", "auth_endpoint")}
        )

    def login(self) -> bool:
        """
        Realiza la autenticación contra la API de DBO y guarda la sesión (JSESSIONID).
        También obtiene automáticamente los datos del perfil actual (author).
        """
        if not self.username or not self.password:
            raise DBOAuthenticationError("No se proporcionaron usuario o contraseña.")

        login_url = f"{self.base_url}{self.auth_endpoint}"
        payload_form = {
            "username": self.username,
            "password": self.password,
            "submit": "Login"
        }

        try:
            resp = self.session.post(
                login_url,
                data=payload_form,
                allow_redirects=False,
                timeout=self.timeout
            )
            
            if resp.status_code in (200, 302):
                location = resp.headers.get("Location", "")
                if "error" in location.lower() or "login" in location.lower():
                    raise DBOAuthenticationError("Credenciales inválidas en DBO.")

                if "JSESSIONID" in self.session.cookies:
                    self._is_authenticated = True
                    logger.info(f"Autenticación exitosa en {login_url}")
                    self._load_current_user()
                    return True

            if resp.status_code == 401:
                raise DBOAuthenticationError("Credenciales inválidas (401 Unauthorized).")

        except requests.RequestException as e:
            raise DBONetworkError(f"Error de red al intentar login en {login_url}: {e}")

        if "JSESSIONID" in self.session.cookies:
            self._is_authenticated = True
            self._load_current_user()
            return True

        raise DBOAuthenticationError(
            f"No se pudo autenticar en DBO ({self.base_url}). Status: {resp.status_code}"
        )

    def _load_current_user(self):
        """Obtiene y guarda el perfil de usuario actual para rellenar el author en los comentarios."""
        try:
            resp = self.session.get(f"{self.base_url}/api/users/current", timeout=self.timeout)
            if resp.status_code == 200:
                data = resp.json()
                self._current_user = data.get("user") or {
                    "id": data.get("id"),
                    "email": data.get("email"),
                    "name": data.get("name"),
                    "creationDate": None,
                    "modificationDate": None
                }
        except Exception as e:
            logger.debug(f"No se pudo cargar /api/users/current: {e}")

    def _get_default_author(self, author: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Devuelve los datos de autor estructurados."""
        return author or self._current_user or {
            "id": 2,
            "email": self.username,
            "name": self.username.split("@")[0] if self.username else "Usuario",
            "creationDate": None,
            "modificationDate": None
        }

    def _post_comment(
        self,
        url: str,
        text: str,
        prefix: str = "",
        author: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Método base para enviar un comentario vía POST a cualquier recurso de DBO."""
        if not self._is_authenticated:
            self.login()

        full_content = f"{prefix} {text}".strip() if prefix else text.strip()
        payload: Dict[str, Any] = {
            "content": full_content,
            "author": self._get_default_author(author)
        }

        last_exception = None
        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.post(
                    url,
                    json=payload,
                    headers={
                        "Content-Type": "application/json",
                        "Accept": "application/json, text/plain, */*"
                    },
                    timeout=self.timeout
                )

                if resp.status_code in (200, 201):
                    try:
                        return resp.json()
                    except ValueError:
                        return {"status": "success", "status_code": resp.status_code}

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al publicar comentario en {url}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red en intento {attempt}/{self.max_retries} para {url}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    def _fetch_comments_page(
        self,
        url: str,
        page: int,
        size: int
    ) -> Dict[str, Any]:
        """Método interno para obtener una página de comentarios con reintentos."""
        if not self._is_authenticated:
            self.login()

        last_exception = None
        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.get(
                    url,
                    params={"page": page, "size": size},
                    headers={"Accept": "application/json, text/plain, */*"},
                    timeout=self.timeout
                )

                if resp.status_code == 200:
                    return resp.json()

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al obtener comentarios de {url}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red al consultar comentarios en {url}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    # ==========================================
    # 📁 EXPEDIENTES (DOCUMENTS)
    # ==========================================
    def expediente_leer_detalle(
        self,
        document_id: Union[int, str]
    ) -> Dict[str, Any]:
        """
        [Expediente_Leer Detalle]
        Obtiene la información completa del expediente (GET /api/documents/{id}).
        """
        if not self._is_authenticated:
            self.login()

        url = f"{self.base_url}/api/documents/{document_id}"
        last_exception = None

        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.get(
                    url,
                    headers={"Accept": "application/json, text/plain, */*"},
                    timeout=self.timeout
                )

                if resp.status_code == 200:
                    return resp.json()

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al obtener documento {document_id}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red al obtener documento {document_id}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    def expediente_actualizar(
        self,
        document_id: Union[int, str],
        data: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        [Expediente_Actualizar]
        Actualiza los datos de un expediente mediante PUT /api/documents/{id}.
        """
        if not self._is_authenticated:
            self.login()

        url = f"{self.base_url}/api/documents/{document_id}"
        last_exception = None

        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.put(
                    url,
                    json=data,
                    headers={
                        "Content-Type": "application/json",
                        "Accept": "application/json, text/plain, */*"
                    },
                    timeout=self.timeout
                )

                if resp.status_code == 200:
                    return resp.json()

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al actualizar documento {document_id}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red al actualizar documento {document_id}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    def expediente_cambiar_estado(
        self,
        document_id: Union[int, str],
        nuevo_estado: str = "WITH_NOTIFICATION_CARD"
    ) -> Dict[str, Any]:
        """
        [Expediente_Cambiar Estado]
        Obtiene los datos actuales del expediente, actualiza el campo 'status' y guarda los cambios con PUT.
        """
        doc = self.expediente_leer_detalle(document_id=document_id)
        doc["status"] = nuevo_estado
        return self.expediente_actualizar(document_id=document_id, data=doc)

    def expediente_crear_comentario(
        self,
        document_id: Union[int, str],
        text: str,
        prefix: str = "*Comentario de Reunion*:",
        author: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        [Expediente_Crear Comentario]
        Publica un comentario en un expediente/documento específico.
        """
        url = f"{self.base_url}/api/documents/{document_id}/comments"
        return self._post_comment(url=url, text=text, prefix=prefix, author=author)

    def expediente_leer_comentarios(
        self,
        document_id: Union[int, str],
        page: int = 0,
        size: int = 20,
        fetch_all: bool = False
    ) -> Union[Dict[str, Any], List[Dict[str, Any]]]:
        """
        [Expediente_Leer Comentario]
        Obtiene los comentarios de un expediente/documento.
        """
        url = f"{self.base_url}/api/documents/{document_id}/comments"
        if fetch_all:
            all_comments: List[Dict[str, Any]] = []
            current_page = 0
            page_size = max(size, 50)
            while True:
                data = self._fetch_comments_page(url, current_page, page_size)
                items = data.get("content", [])
                all_comments.extend(items)
                if data.get("last", True) or not items:
                    break
                current_page += 1
            return all_comments
        return self._fetch_comments_page(url, page, size)

    def expediente_leer_ultimo_comentario(
        self,
        document_id: Union[int, str]
    ) -> Optional[Dict[str, Any]]:
        """
        [Expediente_Leer Ultimo Comentario]
        Obtiene el comentario más reciente de un expediente.
        """
        url = f"{self.base_url}/api/documents/{document_id}/comments"
        data = self._fetch_comments_page(url, page=0, size=1)
        items = data.get("content", [])
        return items[0] if items else None

    # ==========================================
    # ⚙️ SERVICIOS (PERFORMANCES)
    # ==========================================
    def servicio_crear_comentario(
        self,
        service_id: Union[int, str],
        text: str,
        prefix: str = "",
        author: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        [Servicio_Crear Comentario]
        Publica un comentario en un servicio específico.
        """
        url = f"{self.base_url}/api/performances/{service_id}/comments"
        return self._post_comment(url=url, text=text, prefix=prefix, author=author)

    def servicio_leer_comentarios(
        self,
        service_id: Union[int, str],
        page: int = 0,
        size: int = 20,
        fetch_all: bool = False
    ) -> Union[Dict[str, Any], List[Dict[str, Any]]]:
        """
        [Servicio_Leer Comentarios]
        Obtiene los comentarios de un servicio.
        """
        url = f"{self.base_url}/api/performances/{service_id}/comments"
        if fetch_all:
            all_comments: List[Dict[str, Any]] = []
            current_page = 0
            page_size = max(size, 50)
            while True:
                data = self._fetch_comments_page(url, current_page, page_size)
                items = data.get("content", [])
                all_comments.extend(items)
                if data.get("last", True) or not items:
                    break
                current_page += 1
            return all_comments
        return self._fetch_comments_page(url, page, size)

    def servicio_leer_ultimo_comentario(
        self,
        service_id: Union[int, str]
    ) -> Optional[Dict[str, Any]]:
        """
        [Servicio_Leer Ultimo Comentario]
        Obtiene el comentario más reciente de un servicio.
        """
        url = f"{self.base_url}/api/performances/{service_id}/comments"
        data = self._fetch_comments_page(url, page=0, size=1)
        items = data.get("content", [])
        return items[0] if items else None

    # ==========================================
    # 🏢 CLIENTES (CUSTOMERS)
    # ==========================================
    def cliente_leer_detalle(
        self,
        customer_id: Union[int, str]
    ) -> Dict[str, Any]:
        """
        [Cliente_Leer Detalle]
        Obtiene la información completa del cliente (GET /api/customers/{id}).
        """
        if not self._is_authenticated:
            self.login()

        url = f"{self.base_url}/api/customers/{customer_id}"
        last_exception = None

        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.get(
                    url,
                    headers={"Accept": "application/json, text/plain, */*"},
                    timeout=self.timeout
                )

                if resp.status_code == 200:
                    return resp.json()

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al obtener cliente {customer_id}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red al consultar cliente {customer_id}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    def cliente_leer_cuit(
        self,
        customer_id: Union[int, str],
        clean: bool = False
    ) -> Optional[str]:
        """
        [Cliente_Leer CUIT]
        Obtiene el CUIT de un cliente específico.
        """
        cliente = self.cliente_leer_detalle(customer_id=customer_id)
        cuit = cliente.get("cuit")
        if not cuit:
            return None
        
        cuit_str = str(cuit).strip()
        if clean:
            return cuit_str.replace("-", "").replace(" ", "").replace(".", "")
        return cuit_str

    def cliente_leer_credenciales(
        self,
        customer_id: Union[int, str],
        entity: Optional[str] = None
    ) -> Union[List[Dict[str, Any]], Optional[Dict[str, Any]]]:
        """
        [Cliente_Leer Credenciales]
        Obtiene la lista de credenciales asociadas a un cliente (/api/customers/{id}/credentials).
        """
        if not self._is_authenticated:
            self.login()

        url = f"{self.base_url}/api/customers/{customer_id}/credentials"
        last_exception = None

        for attempt in range(1, self.max_retries + 1):
            try:
                resp = self.session.get(
                    url,
                    headers={"Accept": "application/json, text/plain, */*"},
                    timeout=self.timeout
                )

                if resp.status_code == 200:
                    credentials: List[Dict[str, Any]] = resp.json()
                    
                    if entity:
                        entity_lower = entity.lower().strip()
                        for cred in credentials:
                            if entity_lower in cred.get("entity", "").lower():
                                return cred
                        return None
                        
                    return credentials

                if resp.status_code in (401, 403) and attempt == 1:
                    logger.warning("Sesión posiblemente expirada. Reintentando autenticación...")
                    self.login()
                    continue

                raise DBOAPIError(
                    f"Error del servidor al obtener credenciales del cliente {customer_id}: {resp.status_code}",
                    status_code=resp.status_code,
                    response_text=resp.text
                )

            except requests.RequestException as e:
                last_exception = e
                logger.warning(f"Error de red al consultar credenciales del cliente {customer_id}: {e}")
                time.sleep(self.retry_delay)

        raise DBONetworkError(
            f"Fallo persistente de red al conectar con DBO tras {self.max_retries} intentos: {last_exception}"
        )

    # Alias de compatibilidad
    add_comment = expediente_crear_comentario
    get_comments = expediente_leer_comentarios

    def close(self):
        """Cierra la sesión HTTP."""
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
