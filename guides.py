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
### 🧳 Manual de Operaciones: Conciliación de Cobros Viajeros (ME)

Esta herramienta automatiza el cruce de cobros liquidados por viajeros, integrando asientos de caja, bancos y ajustes contables manuales. La lógica está diseñada para limpiar el listado de movimientos que, aunque tengan referencias distintas, ya están compensados financieramente.

---

#### 📂 1. Insumos Requeridos (Archivos Excel)

Debe cargar dos archivos con extensión **.xlsx** que contengan el movimiento analítico de la cuenta:

1.  **Movimientos del Mes Actual:** Exportación del sistema con los nuevos registros del período.
2.  **Saldos del Mes Anterior:** Archivo de "Saldos Abiertos" generado por esta herramienta en el cierre previo.

**Columnas Críticas para el Proceso:**
*   **NIT:** Identificador único del viajero/colaborador.
*   **Asiento:** Prefijos CC (Caja), CB (Bancos) o CG (Ajustes Generales).
*   **Referencia y Fuente:** Campos donde se encuentran los números de recibos y depósitos.
*   **Débito/Crédito Dólar:** Montos en moneda extranjera (la conciliación principal se ejecuta en USD).

---

#### 🧠 2. ¿Cómo funciona la Lógica de Conciliación? (V13)

La herramienta ejecuta un algoritmo de **cuatro fases progresivas** para garantizar que no quede ningún saldo compensado por error:

*   **Fase 0: Depuración de Diferencial:** Identifica y cierra automáticamente líneas de "Ajuste Cambiario" o "Diff", evitando que los céntimos de valoración inflen el reporte de pendientes.
*   **Fase 1: Match de Reversos:** Busca movimientos marcados como "REVERSO". El sistema es capaz de ignorar textos adicionales y encontrar la partida original comparando el NIT y el monto exacto.
*   **Fase 2: Cruce por Inteligencia de Llaves:** 
    *   Analiza los números de recibos/depósitos dentro de las columnas Fuente y Referencia.
    *   Crea un vínculo entre asientos **CC/CG** y **CB** incluso si la información está en columnas cruzadas o si el número fue digitado con sufijos (ej. "12345TI").
*   **Fase 3: Barrido Global por NIT (Cierre Maestro):** 
    *   Es la red de seguridad final. Si un viajero tiene múltiples líneas pendientes que no pudieron emparejarse por número de recibo, el sistema suma el **Saldo Neto Total del NIT**.
    *   Si la suma de débitos y créditos del NIT da **$0.00**, el sistema entiende que la cuenta está saldada y concilia todas las líneas de golpe.

---

#### 🚥 3. Interpretación de Resultados

*   **VIAJERO_[NIT]_[NUMERO]:** Indica que el cruce fue perfecto mediante un identificador de recibo o depósito.
*   **BARRIDO_NETO_NIT_[NIT]:** Indica que se aplicó el cierre maestro; el colaborador no debe dinero al cierre, aunque sus referencias internas no coincidían exactamente.
*   **Tolerancia:** El sistema permite una diferencia de hasta **$0.01** para absorber errores de redondeo derivados de la exportación de Excel.

---

#### 💡 Tips de Uso para el Contador

1.  **NITs Limpios:** Asegúrese de que la columna NIT no tenga caracteres extraños, aunque la herramienta limpia los espacios automáticamente, la uniformidad ayuda a la rapidez del proceso.
2.  **Referencia "TI":** No se preocupe por las referencias que terminan en "TI" (Ajustes de Tesorería); el sistema está programado para ignorar esas letras y extraer solo el número de recibo valioso.
3.  **Ciclo Mensual:** El archivo que hoy descargue como **"Saldos para el Próximo Mes"** debe ser guardado sin modificaciones, ya que será su insumo obligatorio para el proceso del mes siguiente.
""",
    
    "212.05.1108 - Haberes de Clientes": """
        #### 🔎 Lógica de Conciliación Automática (Bolívares - Bs.)

        Manejo de anticipos o saldos a favor de clientes.
        
        1.  **Fase 1: Cruce por NIT:**
            *   Agrupa todos los movimientos de un mismo cliente (NIT). Si la suma de débitos y créditos es cero, se concilia.
        
        2.  **Fase 2: Recuperación por Monto (Sin NIT):**
            *   Si quedan partidas abiertas, busca coincidencias por **Monto Exacto**.
            *   Esto permite cruzar un Débito que tiene el NIT correcto con un Crédito que quizás no tiene NIT (o viceversa), siempre que los montos sean idénticos.
        """,
    "212.07.9001 - CDC - Factoring": """
        #### 🔎 Lógica de Conciliación Automática (Dólares - USD)

        Conciliación de contratos de factoring. El reporte de salida se agrupa por **Proveedor > Contrato**.
        
        1.  **Extracción de Contrato:**
            *   La herramienta analiza la Referencia y la Fuente buscando el código del contrato.
            *   Soporta formatos como: `FQ-xxxx`, `O/Cxxxx`, o números directos (ej: `6016301`) después de la palabra FACTORING.
        
        2.  **Limpieza Automática:**
            *   Elimina automáticamente las líneas de "Diferencia en Cambio".
            
        3.  **Conciliación:**
            *   Agrupa por **NIT** y **Contrato**. Si la suma en Dólares del contrato es cero, se marca como conciliado.
        """,
    "212.05.1005 - Asientos por clasificar": """
        #### 🔎 Lógica de Conciliación Automática (Bolívares - Bs.)

        Esta cuenta transitoria agrupa partidas pendientes de clasificación definitiva. La herramienta aplica una estrategia de 4 fases para limpiarla:
        
        1.  **Limpieza Automática:**
            *   Se detectan y concilian automáticamente las líneas de "Diferencial Cambiario", "Ajustes" o "Diff".
        
        2.  **Cruce por NIT (Fase Principal):**
            *   Agrupa los movimientos por NIT.
            *   Busca pares exactos (1 a 1) que sumen 0.00.
            *   Busca grupos completos (N a N) dentro del mismo NIT que sumen 0.00.
            
        3.  **Cruce Global (Recuperación):**
            *   Busca partidas sueltas que tengan el mismo monto absoluto (cruce por importe) para cerrar casos donde el NIT falte o no coincida.
            
        4.  **Barrido Final:**
            *   Si la suma total de **todos** los movimientos restantes es exactamente **0.00 Bs**, la herramienta asume que son contrapartidas globales y cierra todo el remanente en un solo lote.
        """,
    "115.07.1.002 - Envios en Transito COFERSA": """
### 🚛 Manual de Operaciones: Envíos en Tránsito COFERSA (CRC)

Esta herramienta automatiza la conciliación de la cuenta de tránsitos, utilizando la columna **TIPO** como eje central de los cruces. La lógica está optimizada para manejar grandes volúmenes de datos en **Colones (CRC)**, asegurando un saldo neto de cero en las partidas cerradas.

---

#### 📂 1. Insumos y Columnas Requeridas

Para procesar esta cuenta, debe cargar dos archivos Excel (.xlsx). El sistema identificará automáticamente las siguientes columnas (el radar de la herramienta ignora acentos y diferencia entre mayúsculas/minúsculas):

*   **TIPO:** Es la columna más importante. Contiene los números de embarque (EM... o M...) o categorías de ajuste.
*   **FECHA / ASIENTO / FUENTE:** Columnas de trazabilidad del registro.
*   **REFERENCIA:** Descripción detallada del movimiento.
*   **DÉBITO LOCAL / CRÉDITO LOCAL:** Montos en Colones (Base de la conciliación).
*   **DÉBITO DÓLAR / CRÉDITO DÓLAR:** Montos informativos en USD.

---

#### 🧠 2. ¿Cómo funciona la Lógica de Conciliación? (V16)

La herramienta ya no realiza cruces globales al azar; ahora es **estrictamente jerárquica** dentro de cada grupo de "Tipo":

1.  **Limpieza de Datos:** El sistema ignora filas de totales o celdas vacías provenientes del reporte administrativo (Softland), trabajando solo con asientos contables reales.
2.  **Fase A - Búsqueda de Pares Internos:** Antes de sumar el grupo completo, el sistema revisa cada "Tipo" buscando un Débito y un Crédito que sean **exactamente iguales**. Si los encuentra, los concilia de inmediato (Etiqueta: `PAR_INTERNO`).
3.  **Fase B - Validación de Saldo Neto:** Con los movimientos restantes de cada "Tipo", el sistema realiza una sumatoria algebraica. Si el resultado es **0.00**, cierra todo el bloque (Etiqueta: `GRUPO_NETO`).
4.  **Tolerancia Cero:** Para garantizar la integridad del cierre, la herramienta solo concilia grupos cuyo saldo sea exactamente cero, evitando que queden céntimos huérfanos en la hoja de conciliados.

---

#### 📊 3. Estructura del Reporte de Salida

El archivo generado es dinámico: **solo mostrará las pestañas que contengan datos.**

*   **Agrup. Tipo Abiertas:** Listado de movimientos que tienen un "Tipo" asignado pero que NO sumaron cero (ajustes, reclasificaciones, etc.).
*   **EMB Pendientes:** Exclusivo para números de embarque (**EM** o **M**) que tienen saldo vivo. Incluye totalizadores por cada embarque.
*   **Otros Pendientes:** Movimientos que no tienen nada escrito en la columna "Tipo" y permanecen abiertos.
*   **Especificación:** Hoja principal de auditoría con encabezado oficial de **COFERSA**. Presenta el detalle de saldos abiertos por línea, incluyendo el cálculo de la tasa implícita.
*   **Conciliados:** Histórico de lo cerrado en el proceso. Incluye un totalizador al final para verificar que el **Saldo Neto es 0.00**.

---

#### 💡 Tips para el Éxito en la Conciliación

1.  **Anchos de Columna:** El reporte viene con anchos pre-ajustados para cifras de millones. Si ve `#######`, simplemente ensanche un poco la celda, aunque el sistema ya usa un ancho de 22 para montos.
2.  **Filas Vacías:** No se preocupe si su reporte de Softland trae filas en blanco al final; la herramienta las detecta y las purga automáticamente.
3.  **Hojas Faltantes:** Si el Excel descargado no tiene la hoja "Otros Pendientes", significa que **todos** sus movimientos tenían un Tipo asignado. ¡Es una buena señal de orden contable!
""",
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

# ==============================================================================
# GUÍA PARA EL ANÁLISIS DE PAQUETE CC
# ==============================================================================

GUIA_PAQUETE_CC = """
### 📘 Manual de Operaciones: Análisis de Paquete CC

Esta herramienta clasifica automáticamente los miles de asientos del diario en **Grupos Lógicos** y audita su contenido. Su objetivo es detectar errores antes de la mayorización.

#### 🚥 ¿Cómo leer el reporte? (El Semáforo)

*   ⚪ **Filas Blancas (Conciliado):** El asiento cumple con todas las reglas contables. Está listo para mayorizar.
*   🔴 **Filas Rojas (Incidencia):** El asiento tiene un error o algo inusual. **REQUIERE REVISIÓN MANUAL.**

---

#### 🔍 Qué revisar en cada Grupo (Lógica de Auditoría)

**1. Grupo 1: Acarreos y Fletes Recuperados**
*   **Regla:** La referencia debe contener la palabra "FLETE".
*   **Acción:** Si sale en rojo, verifique por qué se usó la cuenta de fletes sin mencionar fletes en la descripción.

**2. Grupo 2: Diferencial Cambiario**
*   **Qué es:** Ajustes por valoración de moneda (no son cobros reales).
*   **Regla:** Debe contener palabras como `DIFERENCIA`, `CAMBIO`, `TASA`, `AJUSTE`, `DC` o `IVA` (pago diferido).
*   **Ojo:** Si un cobro bancario cae aquí, es un error (debería ir al Grupo 8).

**3. Grupo 3: Notas de Crédito (N/C)**
*   **Estructura Correcta:** Un asiento de N/C por descuento debe tocar dos cuentas: **Descuentos sobre Ventas** + **I.V.A. Débitos Fiscales**.
*   **Error Común (Rojo):** Si falta la cuenta de IVA, la herramienta marcará "Asiento incompleto". Revise si la bonificación fue exenta erróneamente.

**6. Grupo 6: Ingresos Varios (Limpieza)**
*   **Regla del Monto:** Se usa para limpiar centavos o saldos basura.
*   **Límite:** Máximo **$25.00**.
*   **Acción:** Si un asiento supera los $25, saldrá en rojo. Debe reclasificarse o justificarse.

**7. Grupo 7: Devoluciones y Rebajas**
*   **Regla del Monto:** Límite estricto de **$5.00** para ajustes pequeños.
*   **Excepción:** Se permiten montos grandes (millonarios) SOLO SI la referencia indica que es un **TRASLADO**, **CRUCE** o **APLICACIÓN** de saldo entre clientes.
*   **Acción:** Si ve un monto alto en rojo, verifique si falta la palabra "TRASLADO" en la referencia.

**8. Grupo 8: Cobranzas**
*   **Qué es:** Dinero real entrando al banco (TEF, Depósitos) o Recibos de Cobranza.
*   **Validación:** La herramienta agrupa aquí todo lo que toque cuentas de Banco (Mercantil, Banesco, etc.).

**9. Grupo 9: Retenciones (IVA/ISLR)**
*   **Regla:** La referencia debe contener un Número de Comprobante o palabras como `RET` o `IMP`.
*   **Acción:** Si sale en rojo, es porque la referencia está vacía o ilegible.

**11. Grupo 11: Cuentas No Identificadas**
*   **¡ALERTA!** Aquí caen los asientos que usan cuentas contables nuevas o erradas que no están en el sistema.
*   **Acción:** Avise al administrador del sistema para agregar la cuenta al "Directorio de Cuentas" si es correcta.

**13. Grupo 13: Operaciones Reversadas / Anuladas**
*   **Inteligencia Artificial:** La herramienta detectó que hubo un error (ej. una N/C mal hecha) y luego un Reverso que la anuló por el mismo monto.
*   **Estado:** Ambos movimientos se marcan como "Conciliado (Anulado)" y se sacan de los otros grupos para no ensuciar el análisis.

---

#### 💡 Tip de Flujo de Trabajo
Vaya a la hoja **"Listado Correlativo"**. Verá los asientos en orden numérico. Mayorice en lotes hasta que encuentre una **Línea Roja**. Deténgase, corrija ese asiento en el sistema contable, y continúe con el siguiente lote.
"""


GUIA_IMPRENTA = """
### 🖨️ Validación de TXT
Verifica que las facturas del TXT existan en el Libro de Ventas.
"""

GUIA_GENERADOR = """
### ⚙️ Generación de TXT
Crea el archivo de retenciones calculando el prorrateo de montos desde Softland.
"""

GUIA_PENSIONES = """
### 🛡️ Manual de Usuario: Cálculo Ley Protección de Pensiones (9%)

Esta herramienta automatiza el cálculo del aporte del 9%, genera el asiento contable listo para firmar y audita las cifras contra el reporte de RRHH.

---

#### 📂 1. Documentos Requeridos

**A. Mayor Analítico (Contabilidad)**
*   **Fuente:** Sistema Administrativo (Profit/Softland).
*   **Formato:** Excel (`.xlsx`).
*   **Filtros:** Debe descargar el movimiento del mes a declarar.
*   **Cuentas Obligatorias:** El archivo **debe contener** movimientos en:
    *   `7.1.1.01.1.001` (Sueldos y Salarios)
    *   `7.1.1.09.1.003` (Ticket de Alimentación)
*   **Columnas Clave:** Cuenta, Centro de Costo, Débito, Crédito, Fecha.

**B. Resumen de Nómina (RRHH)**
*   **Fuente:** Departamento de Nómina.
*   **Formato:** Excel (`.xlsx`) tipo resumen gerencial.
*   **Pestañas:** El archivo debe tener una pestaña identificada con el **Mes y Año** del cálculo (Ej: "Diciembre 2025" o "Dic-25").
*   **Columnas Requeridas:**
    *   `EMPRESA`: Nombre de la compañía (Febeca, Beval, etc.).
    *   `SALARIOS...`: Monto base de sueldos.
    *   `TICKETS...` o `ALIMENTACION`: Monto base de cestaticket.
    *   `APARTADO`: El monto del impuesto calculado por Nómina (para validar).

---

#### ⚙️ Paso a Paso para una Conciliación Exitosa

1.  **Seleccione la Empresa:** Elija en el menú la compañía a procesar (Ej: QUINCALLA).
2.  **Cargue los Archivos:** Suba el Mayor Contable y el Resumen de Nómina en sus casillas.
3.  **Indique la Tasa:** Ingrese la tasa de cambio (BCV) para que el asiento calcule los Dólares correctamente.
4.  **Ejecute:** Haga clic en "Calcular Impuesto".

#### 🔍 Interpretación de Resultados
*   **✅ Éxito:** Si la Base Contable coincide con la Base de Nómina (Diferencia < 1 Bs), el reporte está listo para imprimir.
*   **⚠️ Descuadre:** Si aparece una alerta amarilla, descargue el Excel y revise la **Hoja 1**. Allí verá una tabla comparativa que le indicará si la diferencia está en los **Salarios** o en los **Tickets**.
"""

GUIA_AJUSTES_USD = """
### 📉 Guía: Ajustes al Balance en USD

Esta herramienta automatiza la valoración de moneda extranjera y reclasificaciones al cierre.

**Insumos Requeridos:**
1.  **Conciliación Bancaria (Excel):** Debe tener la columna "Movimientos en Bancos no Conciliados".
2.  **Balance de Comprobación (Excel/PDF):** El balance general del mes.
3.  **Auxiliares de Viajes:** Reportes de las cuentas 1.1.4.03...
4.  **Reporte Haberes:** Archivo con el "Total de Saldos Negativos" al final.

**Lógica de Ajuste:**
*   **Bancos:** Ajusta según partidas no conciliadas (USD directo o Bs/Tasa Corp).
*   **Saldos Contrarios:** Detecta cuentas negativas y genera el asiento contra su cuenta par.
*   **Haberes:** Incrementa el pasivo según el reporte de saldos negativos.
"""



# ==============================================================================
# GUÍAS GENERALES
# ==============================================================================

GUIA_GENERAL_ESPECIFICACIONES = """
### Guía Práctica: Paso a Paso para Conciliar
1. **Mes Actual:** Arrastre el archivo con los movimientos del mes.
2. **Saldos Anteriores:** Arrastre el archivo de saldos abiertos del mes pasado.
3. **Ejecución:** Haga clic en Iniciar Conciliación y descargue los resultados.
"""

GUIA_COMPLETA_RETENCIONES = "Guía para la auditoría de retenciones IVA e ISLR."
GUIA_PAQUETE_CC = "Manual de clasificación de asientos del diario contable."
GUIA_IMPRENTA = "Validación de archivos TXT contra libros de ventas."
GUIA_GENERADOR = "Generación de archivos de retenciones para GALAC."
GUIA_PENSIONES = "Cálculo del aporte del 9% de la Ley de Pensiones."
GUIA_AJUSTES_USD = "Valoración de activos y pasivos en moneda extranjera al cierre."

GUIA_DEBITO_FISCAL = """
### 📑 Manual de Usuario: Verificación de Débito Fiscal

Esta herramienta realiza una auditoría integral entre la contabilidad (**Softland**) y la información fiscal (**Libro de Ventas de Imprenta**) para asegurar que todo el IVA (Débito Fiscal) facturado esté correctamente registrado.

---

#### 📂 1. Archivos Requeridos

**A. Transacciones de Softland (Diario y Mayor)**
*   **Fuente:** Cuenta `213.04.1001` (IVA Débito Fiscal).
*   **Formato:** Excel (`.xlsx`).
*   **Caso Especial FEBECA-SILLACA:** Debe subir 4 archivos (Diario y Mayor de Febeca + Diario y Mayor de Sillaca). La herramienta los consolidará automáticamente.

**B. Libro de Ventas (Imprenta)**
*   **Estructura:** El sistema asume que los encabezados están en la **Fila 8**.
*   **Columnas Clave:** Se analizan "N de Factura", "N/C", "N/D" e "Impuesto IVA G".

---

#### 🧠 2. Inteligencia de Conciliación

*   **ADN Numérico:** La herramienta limpia documentos (ej: `FAC-000123` -> `123`) y NITs (solo números, ignora J-V-G) para asegurar un match perfecto.
*   **Exclusión de Terceros:** Se descartan automáticamente registros a nombre de **"FEBECA"**, ya que son débitos fiscales a cuenta de terceros.
*   **Filtro de Exentos:** Facturas con IVA 0.00 en el Libro de Ventas son ignoradas para no generar ruidos en la auditoría.
*   **Escudo de Totales:** Se omiten automáticamente las filas verdes de "TOTALES" y los resúmenes de alícuotas del final del libro.

---

#### 🚥 3. ¿Cómo leer el reporte de Incidencias (Hoja 3)?

*   **Listado (Izquierda):** Detalle de diferencias agrupado por Casa (FB/SC) y **Huérfanos** (documentos que están en Imprenta pero nadie ha contabilizado).
*   **Tablas BI (Derecha):** Cuadros de mando ejecutivos que comparan Cantidades y Montos de Softland vs Imprenta.
*   **Validación de Totales:** Los subtotales del listado de incidencias coinciden exactamente con los montos de "Diferencia" de los cuadros de mando.
"""
