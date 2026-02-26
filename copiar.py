"""
STAGING MULTI-AMBIENTE v3.0
- Filtra archivos por ambiente (no mezcla listas)
- Limpia destino antes de copiar (elimina archivos obsoletos)
- Solo copia si el archivo cambió (compara SHA256)
- Log con fecha completa (YYYY-MM-DD HH:MM:SS)
- Matching case-insensitive + búsqueda flexible por nombre similar
- Reporta archivos autorizados no encontrados en origen
- Al finalizar, lanza el proceso de auditoría automáticamente
- Argparse para configuración por CLI
"""

import os
import re
import shutil
import hashlib
import argparse
import subprocess
from datetime import datetime

# ==========================
# ARGUMENTOS CLI
# ==========================
def parse_args():
    parser = argparse.ArgumentParser(
        description="Staging Multi-Ambiente v3.0 - Copia archivos autorizados"
    )
    parser.add_argument("--raiz", default=r"C:\Automatizacion_Excel",
                        help="Carpeta raiz del proyecto")
    parser.add_argument("--no-auditoria", action="store_true",
                        help="No lanzar auditoria al finalizar")
    return parser.parse_args()

_args = parse_args()

# ==========================
# CONFIGURACIÓN
# ==========================
RAIZ             = _args.raiz
RUTA_LISTA       = os.path.join(RAIZ, 'lista_maestra.txt')
ARCHIVO_LOG      = os.path.join(RAIZ, 'LOG_STAGING.txt')
SCRIPT_AUDITORIA = os.path.join(RAIZ, 'procesar.py')
SIN_AUDITORIA    = _args.no_auditoria

AMBIENTES = {
    'CAMARAPROD': {
        'origen':  os.path.join(RAIZ, 'archivos_out', 'INT_CAMARAPROD'),
        'destino': os.path.join(RAIZ, 'PROCESADOS', 'CAMARAPROD'),
        'prefijo': 'camaraprod'
    },
    'CAMARARESP': {
        'origen':  os.path.join(RAIZ, 'archivos_out', 'INT_CAMARARESP'),
        'destino': os.path.join(RAIZ, 'PROCESADOS', 'CAMARARESP'),
        'prefijo': 'camararesp'
    },
    'CAMARATEST': {
        'origen':  os.path.join(RAIZ, 'archivos_out', 'INT_CAMARATEST'),
        'destino': os.path.join(RAIZ, 'PROCESADOS', 'CAMARATEST'),
        'prefijo': 'camaratest'
    }
}


# ==========================
# LOG
# ==========================
def log(msg, nivel="INFO"):
    ts    = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    linea = f"[{ts}] [{nivel}] {msg}"
    print(linea)
    with open(ARCHIVO_LOG, "a", encoding="utf-8") as f:
        f.write(linea + "\n")


# ==========================
# UTILIDADES
# ==========================
def calcular_sha256(ruta):
    sha256 = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(4096), b""):
            sha256.update(bloque)
    return sha256.hexdigest()


def archivos_iguales(src, dst):
    """Retorna True si ambos archivos tienen el mismo SHA256."""
    if not os.path.exists(dst):
        return False
    return calcular_sha256(src) == calcular_sha256(dst)


def construir_mapa_origen(carpeta):
    """
    Construye un diccionario {nombre_lower: nombre_real} para hacer
    búsqueda case-insensitive en la carpeta origen.
    """
    mapa = {}
    if os.path.exists(carpeta):
        for f in os.listdir(carpeta):
            mapa[f.lower()] = f
    return mapa


def _normalizar(s):
    """Normaliza un nombre de archivo para comparación flexible."""
    s = re.sub(r'\.(out|sha256)$', '', s, flags=re.IGNORECASE)
    s = re.sub(r'[\s\-_=,\.]+', '', s)
    return s.lower()


def buscar_en_origen(fname_maestra, mapa_origen):
    """
    Busca el archivo de lista_maestra en el mapa del origen.
    1) Coincidencia exacta (case-insensitive)
    2) Coincidencia normalizada (sin guiones, puntos, extensiones)
    Retorna el nombre real del archivo en origen o None.
    """
    # 1. Exacto case-insensitive
    encontrado = mapa_origen.get(fname_maestra.lower())
    if encontrado:
        return encontrado

    # 2. Normalizado: quitar guiones, espacios, extensiones y comparar
    norm_buscado = _normalizar(fname_maestra)
    for fname_lower, fname_real in mapa_origen.items():
        if _normalizar(fname_lower) == norm_buscado:
            return fname_real

    return None


def detectar_no_listados(mapa_origen, autorizados_lower, prefijo):
    """
    Detecta archivos en la carpeta origen que tienen el prefijo correcto
    y son .out, .sha256 o _out, pero NO están en lista_maestra.
    Retorna lista de nombres reales (del origen) no listados.
    """
    es_dato = lambda f: f.endswith('.out') or f.endswith('.sha256') or f.endswith('_out')
    no_listados = []
    for fname_lower, fname_real in mapa_origen.items():
        if fname_lower.startswith(prefijo) and es_dato(fname_lower):
            if fname_lower not in autorizados_lower:
                no_listados.append(fname_real)
    return sorted(no_listados)



# ==========================
# PROCESO PRINCIPAL
# ==========================
def ejecutar_staging_total():

    # Limpiar log anterior
    if os.path.exists(ARCHIVO_LOG):
        os.remove(ARCHIVO_LOG)

    log("=" * 60)
    log(f"  STAGING MULTI-AMBIENTE v3.0 - INICIO ({datetime.now().strftime('%d/%m/%Y')})")
    log("=" * 60)

    # Cargar lista maestra
    if not os.path.exists(RUTA_LISTA):
        log(f"No existe lista_maestra.txt en {RAIZ}", "ERROR")
        return

    with open(RUTA_LISTA, 'r', encoding='utf-8') as f:
        lista_maestra = [l.strip() for l in f if l.strip()]

    log(f"Lista maestra cargada: {len(lista_maestra)} archivos autorizados")

    resumen_total = {"copiados": 0, "sin_cambios": 0, "eliminados": 0,
                     "no_encontrados": 0, "errores": 0}

    for ambiente, rutas in AMBIENTES.items():
        log(f"\n>>> AMBIENTE: {ambiente}")

        origen  = rutas['origen']
        destino = rutas['destino']
        prefijo = rutas['prefijo']

        # Filtrar lista maestra solo para este ambiente
        autorizados = {f for f in lista_maestra if f.lower().startswith(prefijo)}
        log(f"    Archivos autorizados para este ambiente: {len(autorizados)}")

        if not os.path.exists(origen):
            log(f"    Carpeta origen no encontrada: {origen}", "WARN")
            log(f"    Omitiendo ambiente {ambiente}")
            continue

        os.makedirs(destino, exist_ok=True)

        # --- Limpiar archivos obsoletos del destino ---
        archivos_en_destino = set(os.listdir(destino))
        # Solo eliminar archivos .out, .sha256 y _out (no inventarios ni logs)
        for fname in archivos_en_destino:
            es_dato = (fname.endswith('.out') or fname.endswith('.sha256')
                       or fname.endswith('_out'))
            if es_dato and fname not in autorizados:
                try:
                    os.remove(os.path.join(destino, fname))
                    log(f"    [ELIMINADO] {fname} (ya no está en lista maestra)", "WARN")
                    resumen_total["eliminados"] += 1
                except Exception as e:
                    log(f"    [ERROR] No se pudo eliminar {fname}: {e}", "ERROR")

        # --- Copiar archivos autorizados ---
        mapa_origen       = construir_mapa_origen(origen)  # {nombre_lower: nombre_real}
        autorizados_lower = {f.lower() for f in autorizados}
        copiados          = 0
        sin_cambios       = 0
        no_encontrados    = []
        nuevos_en_lista   = []

        for fname in sorted(autorizados):
            nombre_real = buscar_en_origen(fname, mapa_origen)

            if nombre_real is None:
                log(f"    [FALTA] {fname} — no existe en origen", "WARN")
                no_encontrados.append(fname)
                resumen_total["no_encontrados"] += 1
                continue

            if nombre_real != fname:
                log(f"    [~] {fname} -> encontrado como '{nombre_real}' (nombre diferente)")

            src = os.path.join(origen, nombre_real)
            dst = os.path.join(destino, fname)

            try:
                if archivos_iguales(src, dst):
                    log(f"    [=] {fname} — sin cambios")
                    sin_cambios += 1
                    resumen_total["sin_cambios"] += 1
                else:
                    shutil.copy2(src, dst)
                    hash_val = calcular_sha256(dst)
                    log(f"    [OK] {fname} — copiado (SHA256: {hash_val[:16]}...)")
                    copiados += 1
                    resumen_total["copiados"] += 1
            except Exception as e:
                log(f"    [ERROR] {fname}: {e}", "ERROR")
                resumen_total["errores"] += 1

        # --- Detectar y copiar archivos en origen que NO estan en lista_maestra ---
        no_listados = detectar_no_listados(mapa_origen, autorizados_lower, prefijo)
        if no_listados:
            log(f"    [!] {len(no_listados)} archivos en origen SIN LISTAR — se copian y agregan a lista_maestra:", "WARN")
            for fname_real in no_listados:
                src = os.path.join(origen, fname_real)
                dst = os.path.join(destino, fname_real)
                try:
                    if archivos_iguales(src, dst):
                        log(f"    [=] {fname_real} — sin cambios (no listado)")
                    else:
                        shutil.copy2(src, dst)
                        hash_val = calcular_sha256(dst)
                        log(f"    [NEW] {fname_real} — copiado y AGREGADO a lista maestra (SHA256: {hash_val[:16]}...)", "WARN")
                    nuevos_en_lista.append(fname_real)
                    copiados += 1
                    resumen_total["copiados"] += 1
                except Exception as e:
                    log(f"    [ERROR] {fname_real}: {e}", "ERROR")
                    resumen_total["errores"] += 1

        # --- Inventario del ambiente ---
        inv_path = os.path.join(destino, f"inventario_{prefijo}.txt")
        with open(inv_path, "w", encoding="utf-8") as inv:
            inv.write(f"INVENTARIO {ambiente} - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            inv.write("=" * 60 + "\n\n")
            inv.write(f"Archivos en destino ({len(autorizados) - len(no_encontrados) + len(nuevos_en_lista)}):\n")
            for fname in sorted(list(autorizados) + nuevos_en_lista):
                dst_check = os.path.join(destino, fname)
                if os.path.exists(dst_check):
                    h = calcular_sha256(dst_check)
                    inv.write(f"  {fname} | SHA256: {h}\n")
            if no_encontrados:
                inv.write(f"\nArchivos faltantes en origen ({len(no_encontrados)}):\n")
                for fname in no_encontrados:
                    inv.write(f"  [FALTA] {fname}\n")

        log(f"    Copiados: {copiados} | Sin cambios: {sin_cambios} | Faltantes: {len(no_encontrados)} | Nuevos: {len(nuevos_en_lista)}")
        lista_maestra.extend(nuevos_en_lista)  # acumular para actualizar lista al final

    # --- Actualizar lista_maestra.txt con archivos nuevos detectados ---
    total_originales = sum(1 for l in open(RUTA_LISTA, encoding='utf-8') if l.strip())
    nuevos_total = len(lista_maestra) - total_originales
    if nuevos_total > 0:
        log(f"\n  Actualizando lista_maestra.txt con {nuevos_total} archivos nuevos detectados...", "WARN")
        with open(RUTA_LISTA, "w", encoding="utf-8") as f:
            for fname in sorted(set(lista_maestra)):
                f.write(fname + "\n")
        log(f"  lista_maestra.txt actualizada: {len(set(lista_maestra))} archivos totales.")

    # --- Resumen final ---
    log("\n" + "=" * 60)
    log("  RESUMEN FINAL")
    log("=" * 60)
    log(f"  Archivos copiados:      {resumen_total['copiados']}")
    log(f"  Sin cambios:            {resumen_total['sin_cambios']}")
    log(f"  Eliminados (obsoletos): {resumen_total['eliminados']}")
    log(f"  No encontrados:         {resumen_total['no_encontrados']}")
    log(f"  Nuevos en lista:        {nuevos_total}")
    log(f"  Errores:                {resumen_total['errores']}")

    if resumen_total["errores"] > 0:
        log("\n  ATENCION: Hubo errores durante el proceso. Revisa el log.", "WARN")
        log("  No se ejecutara la auditoria automaticamente.", "WARN")
        return

    # --- Lanzar auditoría automáticamente ---
    if SIN_AUDITORIA:
        log("Opcion --no-auditoria activa: se omite el proceso de auditoria.")
        return

    log("\n" + "=" * 60)
    log("  LANZANDO PROCESO DE AUDITORIA...")
    log("=" * 60)

    if not os.path.exists(SCRIPT_AUDITORIA):
        log(f"Script de auditoria no encontrado: {SCRIPT_AUDITORIA}", "ERROR")
        return

    try:
        resultado = subprocess.run(
            ["python", SCRIPT_AUDITORIA],
            capture_output=True, text=True
        )
        if resultado.returncode == 0:
            log("Auditoria completada exitosamente.")
        else:
            log("La auditoria termino con errores:", "ERROR")
            if resultado.stderr:
                log(resultado.stderr.strip(), "ERROR")
    except Exception as e:
        log(f"No se pudo lanzar la auditoria: {e}", "ERROR")

    log("\n>>> PROCESO COMPLETO <<<")


if __name__ == "__main__":
    ejecutar_staging_total()
    print("\n============================================")
    print("  Proceso Finalizado. Presiona una tecla...")
    input()