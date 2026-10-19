# Excel: Limpieza con Panel de Control  
**Auditoría y control de calidad de datos**

Panel de auditoría en Excel que detecta y cuantifica inconsistencias en un dataset de RR.HH. de 5.000 registros.  
En lugar de revisar miles de filas a ciegas, el panel indica qué controles fallan y cuántos errores hay, para enfocar la limpieza solo donde hace falta.

---

## 📌 Descripción del proyecto

El dataset proviene del repositorio [Pulak-PSR/hr-data-cleaning](https://github.com/Pulak-PSR/hr-data-cleaning) (`hr_data_raw.csv`). Son datos sintéticos de empleados con errores introducidos a propósito para practicar limpieza.

La pregunta que motivó el proyecto fue simple:

> Si cargo datos nuevos en el futuro y aparece un error, ¿cómo me doy cuenta sin revisar manualmente miles de filas?

La respuesta fue construir un **Panel de Control de Auditoría** con fórmulas de Excel que:

- Detectan distintos tipos de inconsistencias.
- Cuentan cuántos registros afectados hay por cada regla.
- Resaltan visualmente (en rojo) cuando el total de errores es mayor que cero.
- Incluyen una guía para localizar y corregir los registros problemáticos.

El panel funciona como un sistema de alerta: si no marca nada, se puede seguir adelante; si marca errores, se va columna por columna a limpiar solo lo necesario.

---

## 📂 Archivos del repositorio

| Archivo | Descripción |
|---------|-------------|
| `01_HR_datos_crudos.csv` | Dataset original sin limpiar (punto de partida). |
| `02_HR_panel_auditoria.xlsx` | Archivo principal: datos en tabla, panel de control, listas de valores permitidos y guía de limpieza. |
| `03_HR_datos_limpios.csv` | Versión final después de las correcciones, lista para análisis o importación. |

---

## 📊 Estructura del dataset

Campos incluidos:

- `employee_id` – identificador único del empleado  
- `first_name` / `last_name` – nombre y apellido  
- `gender` – género  
- `age` – edad  
- `department` – departamento  
- `job_title` – puesto de trabajo  
- `salary` – salario  
- `hire_date` – fecha de contratación  
- `email` – correo electrónico  
- `city` – ciudad  
- `performance_rating` – calificación de desempeño  

---

## 🛠️ Panel de control y reglas de auditoría

El panel usa principalmente `SUMAPRODUCTO` combinada con otras funciones para contar registros que incumplen cada regla.  
Los valores > 0 se resaltan visualmente. Hay un indicador general que se pone en rojo si existe cualquier error.

### 1. Duplicados de `employee_id`
```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*(CONTAR.SI(Datos[employee_id],Datos[employee_id])>1))
```
Cuenta todas las apariciones de un ID repetido (incluyendo la primera). Los vacíos no se consideran.

### 2. Espacios sobrantes
```excel
=SUMAPRODUCTO((Datos[first_name]<>"")*(Datos[first_name]<>ESPACIOS(Datos[first_name])))
```
Detecta espacios al inicio, al final o dobles. Se aplica a columnas de texto relevantes.

### 3. Formato de mayúsculas (Proper Case)
```excel
=SUMAPRODUCTO((Datos[first_name]<>"")*(NO(IGUAL(Datos[first_name],NOMPROPIO(Datos[first_name])))))
```
Compara con el formato que genera `NOMPROPIO`. Útil para nombres y apellidos.

### 4. Valores no permitidos en `job_title`
```excel
=SUMAPRODUCTO((Datos[job_title]<>"")*(CONTAR.SI(Permitidos[job_title],Datos[job_title])=0))
```
Compara contra una lista blanca (`Permitidos`). Es ampliable: solo hay que agregar el nuevo valor correcto a la lista.

### 5. Formato de `employee_id`
```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*((LARGO(Datos[employee_id])<>6)+(IZQUIERDA(Datos[employee_id],1)<>"E")+ESERROR(VALOR(EXTRAE(Datos[employee_id],2,5)))>0))
```
Debe tener exactamente 6 caracteres, empezar con “E” y los 5 restantes ser numéricos.

### 6. Validación básica de email
```excel
=SUMAPRODUCTO((Datos[email]<>"")*((LARGO(Datos[email])-LARGO(SUSTITUIR(Datos[email],"@",""))<>1)+(SI.ERROR(HALLAR(".",Datos[email],HALLAR("@",Datos[email])+1),0)=0)>0))
```
Comprueba exactamente un `@` y la existencia de un punto después del `@`. Es una validación estructural mínima.

### 7. Salarios incorrectos
```excel
=SUMAPRODUCTO((Datos[salary]<>"")*((NO(ESNUMERO(Datos[salary]))+SI.ERROR(Datos[salary]<0,0)+SI.ERROR(Datos[salary]=999999,0))>0))
```
Detecta valores no numéricos, negativos y el valor sospechoso `999999` definido para este dataset.

### 8. Fechas de contratación
```excel
=SUMAPRODUCTO((Datos[hire_date]<>"")*((NO(ESNUMERO(Datos[hire_date]))+(LARGO(TEXTO(Datos[hire_date],"dd/mm/yyyy"))<>10))>0))
```
Verifica que Excel las reconozca como número y que el formato `dd/mm/yyyy` tenga 10 caracteres.  
*(Nota: las fechas originales venían en formato estadounidense y se normalizaron a día-mes-año).*

### 9. Edad numérica
```excel
=SUMAPRODUCTO((Datos[age]<>"")*(NO(ESNUMERO(Datos[age]))))
```
Detecta edades almacenadas como texto.

### 10. Otras validaciones aplicadas
- **Números dentro de nombres/apellidos**: búsqueda de dígitos 0-9 con `HALLAR`.  
- **`performance_rating`**: longitud > 2 caracteres (valores esperados son cortos).  
- **Género**: ajuste de “Non-binary” → “Nonbinary” para evitar interferencia del guion.

---

## 🎯 Objetivo y alcance

Demostrar el uso de Excel para:

1. Definir reglas de calidad por columna.  
2. Detectar y cuantificar inconsistencias de forma automática.  
3. Documentar un procedimiento de limpieza reproducible.  
4. Entregar un dataset limpio listo para análisis o carga en base de datos.

El enfoque es deliberadamente simple y transparente: cualquier persona puede abrir el archivo, ver las fórmulas y entender qué se está controlando.  
No pretende ser una solución enterprise ni usar Power Query/Python; el objetivo es mostrar dominio sólido de limpieza y auditoría con herramientas nativas de Excel.

---

## 📝 Notas

- El dataset es sintético. Algunos comportamientos (como emails repetidos formados por nombre+apellido) no ocurren en datos reales y se trataron en consecuencia.  
- Las reglas son específicas de este conjunto. Para otros datasets hay que adaptar o agregar controles.  
- El panel no corrige automáticamente: solo alerta. La corrección se hace de forma guiada y documentada.

---

**Herramientas**: Excel (fórmulas, tablas, formato condicional)  

