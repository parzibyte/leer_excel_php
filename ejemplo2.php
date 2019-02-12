<?php
/**
 * Demostrar lectura de hoja de cálculo o archivo
 * de Excel con PHPSpreadSheet: leer determinada celda
 * por número de columna y fila 
 *
 * @author parzibyte
 */
# Cargar librerias y cosas necesarias
require_once "vendor/autoload.php";

# Indicar que usaremos el IOFactory
use PhpOffice\PhpSpreadsheet\IOFactory;

# Recomiendo poner la ruta absoluta si no está junto al script
# Nota: no necesariamente tiene que tener la extensión XLSX
$rutaArchivo = "LibroParaLeerConPHP.xlsx";
$documento = IOFactory::load($rutaArchivo);

# Recuerda que un documento puede tener múltiples hojas
# obtener conteo e iterar
$totalDeHojas = $documento->getSheetCount();

# Iterar hoja por hoja
for ($indiceHoja = 0; $indiceHoja < $totalDeHojas; $indiceHoja++) {

    # Obtener hoja en el índice que vaya del ciclo
    $hojaActual = $documento->getSheet($indiceHoja);
    echo "<h3>Vamos en la hoja con índice $indiceHoja</h3>";

    # Nota: las columnas y filas comienzan en 1, no en 0
    $columna = 1;
    $fila = 1;

    # Lo que hay en 1, 1
    $celda = $hojaActual->getCellByColumnAndRow($columna, $fila);
    # El valor, así como está en el documento
    $valorRaw = $celda->getValue();

    # Formateado por ejemplo como dinero o con decimales
    $valorFormateado = $celda->getFormattedValue();

    # Si es una fórmula y necesitamos su valor, llamamos a:
    $valorCalculado = $celda->getCalculatedValue();

    # Imprimir
    echo "En <strong>$columna, $fila</strong> tenemos el valor <strong>$valorRaw</strong>. ";
    echo "Formateado es: <strong>$valorFormateado</strong>. ";
    echo "Calculado es: <strong>$valorCalculado</strong><br><br>";

}
