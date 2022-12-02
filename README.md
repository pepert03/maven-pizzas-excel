Reporte ejecutivo en Excel
=================

## Introducción
Instalar librerías de Python necesarias:
```
pip install -r requirements.txt
```
Todos los archivos necesarios se encuentran en la carpeta `data`. Entre ellos se encuentra ```predicción.csv``` obtenido en 
[maven-pizzas-xml](https://github.com/pepert03/maven-pizzas-xml) y ```ganancias_mensuales.json``` obtenido en [maven-pizzas-pdf](https://github.com/pepert03/maven-pizzas-pdf).

## Ejecución
El archivo ```pizzas_to_excel.py``` genera un ```reporte_ventas.xlsx``` con 3 worksheets:
* Reporte de ventas, con un grafico de barras de las ventas por mes.
* Reporte de ingredientes, con un grafico de la meddia de ingredientes usados por mes.
* Reporte de pedidos, con un grafico de pedidos por dia de la semana.