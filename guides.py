# guides.py

# ==============================================================================
# GUÍA GENERAL PARA LA HERRAMIENTA DE ESPECIFICACIONES
# ==============================================================================

GUIA_GENERAL_ESPECIFICACIONES = """
### Guía Práctica: Paso a Paso para Conciliar

Siga estos 3 pasos para garantizar una conciliación exitosa y sin errores.

---

#### **Paso 1: Preparación de los 2 Archivos de Entrada**

La calidad de la conciliación depende de la correcta preparación de los datos. Asegúrese de que sus archivos `.xlsx` cumplan con lo siguiente:

**1. 📂 Movimientos del Mes Actual:**
*   Contiene todas las transacciones del período que está cerrando.
*   Debe estar en la **primera hoja** del archivo Excel.

**2. 📂 Saldos del Mes Anterior:**
*   Contiene todas las partidas que quedaron abiertas (pendientes) del ciclo de conciliación anterior.
*   **ACCIÓN CRÍTICA:** Para el primer uso, este archivo puede ser su reporte de saldos abiertos. Para los meses siguientes, **debe usar el archivo CSV (`saldos_para_proximo_mes.csv`)** que genera esta misma herramienta al finalizar cada proceso.

Ambos archivos deben contener las **columnas esenciales** que se listan en el recuadro azul informativo justo debajo de esta guía.

---

#### **Paso 2: Carga y Ejecución**

1.  **Seleccione la Empresa (Casa)** y la **Cuenta Contable** que desea procesar.
2.  Arrastre y suelte (o busque) los dos archivos en sus respectivas cajas de carga.
3.  Haga clic en el botón **"▶️ Iniciar Conciliación"**.

---

#### **Paso 3: Descarga y Continuidad del Ciclo**

1.  Una vez finalizado, descargue el **Reporte Completo (Excel)** para su análisis y archivo.
2.  **MUY IMPORTANTE:** Descargue los **Saldos para Próximo Mes (CSV)**. Este archivo es su nuevo punto de partida y deberá usarlo como el archivo de "Saldos anteriores" en la próxima conciliación de esta misma cuenta.
"""

# ==============================================================================
# DICCIONARIO DE GUÍAS ESPECÍFICAS POR CUENTA
# ==============================================================================

LOGICA_POR_CUENTA = {
    "111.04.1001 - Fondos en Tránsito": """
        #### 🔎 Lógica de Conciliación Automática (Bolívares - Bs.)
        
        Esta cuenta tiene una lógica de conciliación muy detallada que se ejecuta en múltiples fases, buscando agrupar y anular movimientos que se corresponden entre sí.
        
        1.  **Conciliación Inmediata:**
            *   Todos los movimientos cuya referencia contenga `DIFERENCIA EN CAMBIO`, `DIF. CAMBIO` o `AJUSTE` se concilian automáticamente.
        
        2.  **Análisis por Categoría de Referencia:**
            *   La herramienta primero clasifica cada movimiento en grupos según palabras clave en su referencia: **SILLACA**, **NOTA DE DEBITO/CREDITO**, **BANCO A BANCO**, **REMESA**, etc.
            *   Dentro de cada uno de estos grupos, intenta conciliar de la forma más específica a la más general:
                *   Busca **pares exactos** (un débito y un crédito) que se anulen (sumen 0) y compartan la misma referencia.
                *   Busca **pares aproximados** que se anulen dentro de una pequeña tolerancia.
                *   Busca **grupos de movimientos** (N vs N) que compartan la misma **Fecha** o **Referencia** y cuya suma total sea cero.
                *   Si al final de analizar una categoría todos los movimientos restantes suman cero, los concilia como un **lote**.
        
        3.  **Búsqueda Global Final:**
            *   Después de analizar por categorías, la herramienta revisa **todos los movimientos pendientes** y busca pares o grupos que compartan la misma referencia literal (ej: un número de transferencia) y se anulen entre sí.
        """,

    "111.04.6001 - Fondos por Depositar - ME": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        Esta cuenta se concilia en **Dólares (USD)** y sigue una estrategia de múltiples pasos para encontrar contrapartidas.
        
        1.  **Conciliación Inmediata:**
            *   Al igual que en Fondos en Tránsito, las `DIFERENCIA EN CAMBIO` y `AJUSTE` se concilian de inmediato.
        
        2.  **Grupos por Referencia:**
            *   Busca todos los movimientos (2 o más) que compartan **exactamente la misma referencia normalizada** (ej: "BANCARIZACIONLOTE5") y los concilia si su suma total en USD es cero (o casi cero).
        
        3.  **Pares por Monto Exacto:**
            *   Busca en todos los movimientos pendientes un débito y un crédito que tengan el **mismo valor absoluto**. Por ejemplo, un débito de `$500.00` se conciliará con un crédito de `-$500.00`, sin importar la referencia. Se da prioridad a los movimientos tipo `BANCO A BANCO`.
        
        4.  **Grupos Complejos (1 vs N o N vs 1):**
            *   Realiza una búsqueda avanzada para encontrar situaciones donde un movimiento grande es la contrapartida de varios pequeños. Por ejemplo, busca si **1 débito** se anula con la suma de **2 créditos**, o si **2 débitos** se anulan con la suma de **1 crédito**.
            
        5.  **Conciliación Final por Lote:**
            *   Si después de todos los pasos anteriores, la **suma total de todos los movimientos pendientes** es cero (o casi cero), los concilia a todos como un lote de cierre.
        """,
        
    "212.07.6009 - Devoluciones a Proveedores": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        La lógica para esta cuenta es muy específica y se basa en cruzar la información de las devoluciones con sus notas de crédito correspondientes.
        
        1.  **Generación de Llaves de Cruce:**
            *   🔑 **Llave 1 (Proveedor):** Se utiliza el **NIT/RIF** del proveedor como identificador único.
            *   🔑 **Llave 2 (Comprobante):** Para las devoluciones (débitos), se usa el dato de la columna `Fuente`. Para las notas de crédito (créditos), se extrae el número de comprobante (ej: `COMP-12345`) de la columna `Referencia`.
        
        2.  **Conciliación por Grupo:**
            *   La herramienta agrupa todos los movimientos que compartan **el mismo Proveedor Y el mismo número de Comprobante**.
            *   Si la suma en **Dólares (USD)** de uno de estos grupos es cero (o casi cero), todos los movimientos dentro de ese grupo se marcan como conciliados.
        """,
        
    "114.03.1002 - Cuentas de viajes - anticipos de gastos": """
        #### 🔎 Lógica de Conciliación Automática (Bolívares - Bs.)

        Esta cuenta busca anular los anticipos de viaje con sus respectivas legalizaciones, utilizando el NIT del colaborador como ancla principal.
        
        1.  **Generación de Clave:**
            *   🔑 Se utiliza el **NIT/RIF** del colaborador o proveedor como la clave principal de agrupación.
        
        2.  **Búsqueda de Pares Exactos:**
            *   Para un mismo NIT, la herramienta busca un débito y un crédito que tengan el **mismo valor absoluto exacto**. Por ejemplo, un anticipo de `5,000.00 Bs` se conciliará con una legalización de `-5,000.00 Bs` para el mismo colaborador.
            
        3.  **Búsqueda de Grupos por Saldo Cero:**
            *   Si no encuentra pares exactos, la herramienta agrupa **todos los movimientos pendientes de un mismo NIT**.
            *   Si la suma total en **Bolívares (Bs.)** de todos esos movimientos es cero (o casi cero), los concilia a todos como un grupo.
            *   También intenta buscar sub-grupos más pequeños dentro de los movimientos de un NIT que puedan sumar cero.
        """,
        
    "114.02.6006 - Deudores Empleados - Otros (ME)": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        La lógica de esta cuenta es directa y se enfoca en verificar el saldo final de cada empleado en moneda extranjera.
        
        1.  **Generación de Clave:**
            *   🔑 Se utiliza el **NIT/RIF** del empleado como el identificador único para agrupar todos sus movimientos.
        
        2.  **Conciliación por Saldo Total del Empleado:**
            *   La herramienta calcula el saldo total en **Dólares (USD)** sumando todos los débitos y créditos para cada empleado.
            *   Si el saldo final de un empleado es **cero (o un valor muy cercano a cero)**, todos sus movimientos se marcan como conciliados. La lógica asume que la cuenta del empleado está saldada.
        """
}
