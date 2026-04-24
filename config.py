"""
Configuración de conexión Oracle y variables de entorno.
"""
import os
from pathlib import Path
from dataclasses import dataclass


@dataclass
class OracleConfig:
    """Configuración para conexión Oracle RAC."""
    user: str
    password: str
    targets: list[tuple[str, str, str]]  # (host, port, sid)
    jdbc_jar: Path
    
    @classmethod
    def from_env(cls) -> 'OracleConfig':
        """Carga configuración desde variables de entorno."""
        jar_path = Path(os.environ.get("ORACLE_JDBC_JAR", "jdbc/ojdbc8.jar")).expanduser()
        if not jar_path.exists():
            raise RuntimeError(f"Driver JDBC no encontrado: {jar_path}")
        
        targets_str = os.environ.get("ORACLE_TARGETS", "")
        if not targets_str:
            raise RuntimeError("ORACLE_TARGETS no configurado")
        
        targets = []
        for target in targets_str.split(","):
            parts = target.strip().split(":")
            if len(parts) == 3:
                targets.append((parts[0], parts[1], parts[2]))
        
        if not targets:
            raise RuntimeError("ORACLE_TARGETS formato inválido (host:port:sid,...)")
        
        return cls(
            user=os.environ.get("ORACLE_USER", ""),
            password=os.environ.get("ORACLE_PASSWORD", ""),
            targets=targets,
            jdbc_jar=jar_path
        )


def project_root() -> Path:
    """Devuelve la raíz del proyecto."""
    return Path(__file__).resolve().parent


def get_template_excel_path() -> Path:
    """
    Devuelve la ruta de la plantilla oficial de homologación.
    """
    path = project_root() / "resources" / "templates" / "plantilla_homologacion_items_issfa.xlsx"
    if not path.exists():
        raise RuntimeError(
            f"No se encontró la plantilla oficial en:\n{path}"
        )
    return path


def get_oracle_targets() -> list[tuple[str, str, str]]:
    """Obtiene lista de nodos Oracle desde variable de entorno."""
    targets_str = os.environ.get("ORACLE_TARGETS", "")
    targets = []
    for target in targets_str.split(","):
        parts = target.strip().split(":")
        if len(parts) == 3:
            targets.append((parts[0], parts[1], parts[2]))
    return targets
