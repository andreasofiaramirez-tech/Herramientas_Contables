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
*   **ACCIÓN CRÍTICA:** Para el primer uso, este archivo puede ser su reporte de saldos abiertos. Para los meses siguientes, **debe usar el archivo Excel (`saldos_para_proximo_mes.xls`)** que genera esta misma herramienta al finalizar cada proceso.

Ambos archivos deben contener las **columnas esenciales** que se listan en el recuadro azul informativo justo debajo de esta guía.

---

#### **Paso 2: Carga y Ejecución**

1.  **Seleccione la Empresa (Casa)** y la **Cuenta Contable** que desea procesar.
2.  Arrastre y suelte (o busque) los dos archivos en sus respectivas cajas de carga.
3.  Haga clic en el botón **"▶️ Iniciar Conciliación"**.

---

#### **Paso 3: Descarga y Continuidad del Ciclo**

1.  Una vez finalizado, descargue el **Reporte Completo (Excel)** para su análisis y archivo.
2.  **MUY IMPORTANTE:** Descargue los **Saldos para Próximo Mes (excel)**. Este archivo es su nuevo punto de partida y deberá usarlo como el archivo de "Saldos anteriores" en la próxima conciliación de esta misma cuenta.
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
        """,
    
    "111.04.6003 - Fondos por Depositar - Cobros Viajeros - ME": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        Esta cuenta gestiona la liquidación de cobros de viajeros, enfocándose en cruzar la cobranza (CC) con su depósito o registro bancario (CB).

        1.  **Agrupación Principal por NIT:**
            *   🔑 El **NIT** del viajero es la clave fundamental. La herramienta nunca mezclará movimientos de clientes diferentes.

        2.  **Fase 1: Detección Inteligente de Reversos:**
            *   La herramienta busca movimientos marcados como **"REVERSO"**.
            *   Utiliza una lógica de **coincidencia parcial**: si un reverso tiene la referencia "REV-12345" y existe un movimiento original "12345" (o viceversa) para el mismo NIT, y sus montos se anulan, los concilia automáticamente.

        3.  **Fase 2: Cruce Estándar (N-a-N):**
            *   Para el resto de movimientos, la herramienta construye una **"Clave de Vínculo"** extrayendo solo los números de la Referencia o la Fuente, dependiendo del tipo de asiento (CC vs CB).
            *   Agrupa todos los movimientos de un mismo NIT que compartan ese número de vínculo (ej. un número de planilla o recibo).
            *   Si la suma total del grupo es **cero (0.00 USD)**, se marcan todos como conciliados.
        """,
    "212.05.1108 - Haberes de Clientes": """
        #### 🔎 Lógica de Conciliación Automática (Bolívares - Bs.)

        Esta cuenta maneja los anticipos o saldos a favor de clientes.
        
        1.  **Fase 1: Cruce por NIT:**
            *   Agrupa todos los movimientos de un mismo cliente (NIT). Si la suma de débitos y créditos es cero, se concilia.
        
        2.  **Fase 2: Recuperación por Monto (Sin NIT):**
            *   Si quedan partidas abiertas, busca coincidencias por **Monto Exacto**.
            *   Esto permite cruzar un Débito que tiene el NIT correcto con un Crédito que quizás no tiene NIT (o viceversa), siempre que los montos sean idénticos.
        """,
    "212.07.9001 - CDC - Factoring": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        Conciliación de contratos de factoring basada en la referencia del documento.
        
        1.  **Extracción de Contrato:**
            *   Busca en la referencia el patrón `FACTORING [CODIGO] $`.
            *   Extrae el código que se encuentra entre la palabra "FACTORING" y el signo de dólar.
        
        2.  **Conciliación por Grupo:**
            *   Agrupa por **NIT** y **Contrato**.
            *   Si la suma en Dólares de ese contrato es cero, se marca como conciliado.
        """
}

# ==============================================================================
# GUÍA PARA LA HERRAMIENTA DE RETENCIONES
# ==============================================================================

GUIA_COMPLETA_RETENCIONES = """
### Guía Práctica: Paso a Paso para el Uso Correcto

Siga estos 4 pasos para garantizar una auditoría exitosa y sin errores.

---

#### **Paso 1: Preparación de los 5 Archivos de Entrada**

La calidad de la auditoría depende de la correcta preparación de los datos. Asegúrese de que sus archivos `.xlsx` cumplan con lo siguiente:

**1. 📂 Relacion_Retenciones_CP.xlsx (Su archivo de trabajo)**
*   **Formato:** Los encabezados de la tabla deben estar **exactamente en la fila 5**.
*   **Columnas Esenciales Requeridas:**
    - `Asiento Contable`
    - `Proveedor` (Debe contener el RIF del proveedor)
    - `Tipo`
    - `Fecha`
    - `Número` (El número de comprobante de retención)
    - `Monto`
    - `Aplicación` (Aquí se busca el número de factura)
    - `Subtipo` (Debe contener 'IVA', 'ISLR' o 'MUNICIPAL')

**2. 📂 Transacciones_Diario_CG.xlsx (Su reporte del diario contable)**
*   **ACCIÓN CRÍTICA:** Antes de exportar, **filtre el diario contable** para incluir únicamente los asientos cuyo rango de fechas coincida con el de su archivo CP. Esto acelera el proceso y evita falsos negativos.
*   **Columnas Esenciales Requeridas:**
    - `ASIENTO`
    - `CUENTACONTABLE`
    - `DEBITOVES` (o un nombre similar como DÉBITO, DEBEVESDÉBITO)
    - `CREDITOVES` (o un nombre similar como CRÉDITO)

**3, 4 y 5. 📂 Archivos de GALAC (IVA, ISLR, Municipales)**
*   Estos deben ser los reportes oficiales generados por el sistema, sin modificaciones. La herramienta está programada para leer su estructura nativa.

---

#### **Paso 2: Carga de Archivos en la Herramienta**

1.  Arrastre y suelte (o busque) cada uno de los 5 archivos en su respectiva caja de carga en la interfaz.
2.  La aplicación reconocerá los archivos y activará el botón de inicio.

---

#### **Paso 3: Ejecución y Descarga del Reporte**

1.  Haga clic en el botón **"▶️ Iniciar Auditoría de Retenciones"**.
2.  Espere mientras la herramienta procesa y concilia todos los registros.
3.  Una vez finalizado, aparecerá el botón **"⬇️ Descargar Reporte de Auditoría (Excel)"**. Haga clic para obtener su archivo de resultados.

---

#### **Paso 4: Interpretación de los Resultados en el Excel**

El reporte de Excel generado tiene dos columnas clave que resumen el estado de cada registro:

*   **`Cp Vs Galac`**: Le dice si su registro de CP coincide con la fuente oficial.
    - **`Sí`**: ¡Perfecto! El registro de CP coincide con GALAC.
    - **`Anulado`**: El registro fue marcado como anulado en su CP.
    - **`Comprobante no encontrado`**: El número de comprobante, para ese RIF, no existe en el reporte de GALAC. Verifique el número y el RIF.
    - **`Error de Subtipo`**: El registro fue encontrado, pero en un tipo de retención diferente (ej: se declaró como IVA pero se encontró en ISLR).

*   **`Validacion CG`**: Una vez validado con GALAC, se verifica contra el diario contable.
    - **`Conciliado en CG`**: ¡Éxito! El asiento, la cuenta contable y el monto son correctos en el diario.
    - **`Asiento no encontrado en CG`**: El número de asiento de su CP no existe en el archivo del diario que subió.
    - **`Cuenta Contable no coincide`**: El asiento se registró en una cuenta que no corresponde al tipo de retención.
    - **`Monto no coincide`**: El monto del débito/crédito en el diario no coincide con el monto de su CP.

💡 **Un registro está 100% conciliado solo si ambas columnas muestran un estado exitoso.**

---
### Análisis Detallado: ¿Cómo Funciona la Lógica de Conciliación?

La herramienta realiza una auditoría automática en dos fases cruciales:

#### **Fase 1: Validación Cruzada (CP vs. GALAC)**
Se asegura que lo preparado en la **Contabilidad Preparada (CP)** coincida con la fuente oficial **GALAC**. La lógica varía según el tipo de retención (IVA, ISLR, Municipal) buscando siempre una combinación de **RIF, Comprobante, Factura y Monto**.

#### **Fase 2: Verificación Contable Final (CP vs. CG)**
Una vez validado contra GALAC, se asegura que el registro fue correctamente asentado en la **Contabilidad General (CG)**, usando el **Número de Asiento** como llave para verificar la **Cuenta Contable** y el **Monto** correctos.
"""
