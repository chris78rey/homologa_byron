"""
Conexión Oracle y operaciones de base de datos.
"""
import os
from pathlib import Path
from typing import Optional
from contextlib import contextmanager
import jaydebeapi
from config import get_oracle_targets


# Configuración del driver JDBC
JAR_PATH = Path(os.environ.get("ORACLE_JDBC_JAR", "jdbc/ojdbc8.jar")).expanduser()
DRIVER = "oracle.jdbc.OracleDriver"


class DatabaseError(Exception):
    """Excepción para errores de base de datos."""
    pass


def _build_jdbc_url(host: str, port: str, sid: str) -> str:
    """Construye URL JDBC para Oracle."""
    return f"jdbc:oracle:thin:@{host}:{port}:{sid}"


@contextmanager
def oracle_connection(user: str, password: str):
    """
    Conexión a Oracle con failover automático entre nodos RAC.
    """
    targets = get_oracle_targets()
    if not targets:
        raise DatabaseError("No hay nodos Oracle configurados")
    
    last_error = None
    for host, port, sid in targets:
        url = _build_jdbc_url(host, port, sid)
        try:
            conn = jaydebeapi.connect(
                DRIVER, 
                url, 
                [user, password], 
                jars=[str(JAR_PATH)]
            )
            try:
                conn.jvm.startJvm(jars=[str(JAR_PATH)])
            except Exception:
                pass  # JVM ya iniciado
            yield conn
            return
        except Exception as e:
            last_error = e
            continue
    
    raise DatabaseError(f"No se pudo conectar a ningún nodo Oracle: {last_error}")


class OracleConnection:
    """Gestor de conexión Oracle con reconexión automática."""
    
    def __init__(self, user: str, password: str):
        self.user = user
        self.password = password
        self._conn = None
        self._cursor = None
    
    def connect(self) -> bool:
        """Establece conexión a Oracle."""
        try:
            targets = get_oracle_targets()
            for host, port, sid in targets:
                url = _build_jdbc_url(host, port, sid)
                try:
                    self._conn = jaydebeapi.connect(
                        DRIVER, 
                        url, 
                        [self.user, self.password], 
                        jars=[str(JAR_PATH)]
                    )
                    self._cursor = self._conn.cursor()
                    return True
                except Exception:
                    continue
            return False
        except Exception:
            return False
    
    def disconnect(self):
        """Cierra conexión."""
        if self._cursor:
            try:
                self._cursor.close()
            except Exception:
                pass
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
        self._cursor = None
        self._conn = None
    
    def execute(self, sql: str, params: tuple = None) -> list:
        """Ejecuta consulta SQL y retorna resultados."""
        if not self._cursor:
            raise DatabaseError("No hay conexión activa")
        try:
            if params:
                self._cursor.execute(sql, params)
            else:
                self._cursor.execute(sql)
            return self._cursor.fetchall()
        except Exception as e:
            raise DatabaseError(f"Error executing SQL: {e}")
    
    def execute_many(self, sql: str, params_list: list):
        """Ejecuta SQL con múltiples parámetros."""
        if not self._cursor:
            raise DatabaseError("No hay conexión activa")
        try:
            self._cursor.executemany(sql, params_list)
        except Exception as e:
            raise DatabaseError(f"Error executing batch: {e}")
    
    def commit(self):
        """Confirma transacción."""
        if self._conn:
            self._conn.commit()
    
    def rollback(self):
        """Revierte transacción."""
        if self._conn:
            self._conn.rollback()
        
    def rollback_force(self):
        """Fuerza rollback real."""
        if self._conn:
            self._conn.rollback()
    
    @property
    def cursor(self):
        return self._cursor
    
    @property
    def connection(self):
        return self._conn


# ============================================================================
# CONSULTAS SQL PARA HOMOLOGACIÓN
# ============================================================================

SQL_CHECK_DETALLE_EXISTS = """
    SELECT ID_ITISF, CODIGO, DESCRIPCION, TIPO
    FROM SIS.ITEMS_ISSFA_DETALLE
    WHERE ID_ITISF = :1 AND CODIGO = :2
"""

SQL_INSERT_DETALLE = """
    INSERT INTO SIS.ITEMS_ISSFA_DETALLE (ID_ITISF, CODIGO, DESCRIPCION, TIPO)
    VALUES (:1, :2, :3, :4)
"""

SQL_CHECK_ITEMS_EXISTS = """
    SELECT TIPO, SBS_SCC_CODIGO, SBS_CODIGO, CODIGO, DESCRIPCION, UNIDAD
    FROM SIS.ITEMS
    WHERE CODIGO = :1
"""

SQL_CHECK_EQUIVALENCIA_EXISTS = """
    SELECT TIPO, SBS_SCC_CODIGO, SBS_CODIGO, CODIGO, ID_ITISF, CODIGO_ISSFA
    FROM SIS.EQUIVALENCIAS_ITEMS_ISSFA
    WHERE TIPO = :1 AND SBS_SCC_CODIGO = :2 AND SBS_CODIGO = :3
      AND CODIGO = :4 AND ID_ITISF = :5 AND CODIGO_ISSFA = :6
"""

SQL_INSERT_EQUIVALENCIA = """
    INSERT INTO SIS.EQUIVALENCIAS_ITEMS_ISSFA
        (TIPO, SBS_SCC_CODIGO, SBS_CODIGO, CODIGO, ID_ITISF, CODIGO_ISSFA)
    VALUES (:1, :2, :3, :4, :5, :6)
"""

SQL_COUNT_DETALLE = "SELECT COUNT(*) FROM SIS.ITEMS_ISSFA_DETALLE"
SQL_COUNT_EQUIVALENCIAS = "SELECT COUNT(*) FROM SIS.EQUIVALENCIAS_ITEMS_ISSFA"

SQL_BACKUP_DETALLE = """
    CREATE TABLE SIS.BKP_ITEMS_ISSFA_DETALLE_{date} AS
    SELECT * FROM SIS.ITEMS_ISSFA_DETALLE
"""

SQL_BACKUP_EQUIVALENCIAS = """
    CREATE TABLE SIS.BKP_EQUIVALENCIAS_ITEMS_ISSFA_{date} AS
    SELECT * FROM SIS.EQUIVALENCIAS_ITEMS_ISSFA
"""

SQL_PANIC_RESTORE_EQUIVALENCIAS = """
    DELETE FROM SIS.EQUIVALENCIAS_ITEMS_ISSFA
"""

SQL_PANIC_RESTORE_DETALLE = """
    DELETE FROM SIS.ITEMS_ISSFA_DETALLE
"""
