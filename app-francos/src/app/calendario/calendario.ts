import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './calendario.html',
  styleUrls: ['./calendario.css']
})
export class Calendario {
  nombreBusqueda: string = '';
  datosExcel: any[] = []; // Aquí se guardará toda la tabla del Excel
  resultados: any[] = []; // Aquí los francos de la persona buscada
  cargando: boolean = false;
  columnas: string[] = []; // Para saber qué columnas tiene tu Excel

  // 1. Leer el archivo Excel subido
  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.cargando = true;
    const fileReader = new FileReader();
    
    fileReader.onload = (e: any) => {
      const arrayBuffer = e.target.result;
      const data = new Uint8Array(arrayBuffer);
      
      // Leer el libro de trabajo (Workbook)
      const workbook = XLSX.read(data, { type: 'array' });
      
      // Tomamos la primera hoja del Excel
      const nombreHoja = workbook.SheetNames[0];
      const hoja = workbook.Sheets[nombreHoja];
      
      // Convertimos la hoja a un formato JSON (Arreglo de objetos)
      this.datosExcel = XLSX.utils.sheet_to_json(hoja);

      if (this.datosExcel.length > 0) {
        // Guardamos los nombres de las columnas para poder mostrarlas en la tabla de la app
        this.columnas = Object.keys(this.datosExcel[0]);
        alert('¡Excel cargado y procesado con éxito!');
      } else {
        alert('El archivo Excel está vacío.');
      }
      
      this.cargando = false;
    };

    fileReader.readAsArrayBuffer(file);
  }

  // 2. Filtrar la información por el nombre del empleado
  buscarFrancos() {
    if (!this.nombreBusqueda.trim()) {
      alert('Por favor, ingresa un nombre.');
      return;
    }

    const termino = this.nombreBusqueda.toLowerCase().trim();

    // Buscamos en todas las filas del Excel. 
    // Como no sé exactamente cómo se llama tu columna de nombres, 
    // este código busca el término en cualquier celda de la fila.
    this.resultados = this.datosExcel.filter(fila => {
      return Object.values(fila).some(valor => 
        String(valor).toLowerCase().includes(termino)
      );
    });
  }
}