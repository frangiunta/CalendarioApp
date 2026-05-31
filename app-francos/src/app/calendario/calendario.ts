import { Component, OnInit, ViewEncapsulation } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

// Importaciones de Angular Material
import { MatCardModule } from '@angular/material/card';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatCalendarCellClassFunction, MatDatepickerModule } from '@angular/material/datepicker';
import { MatNativeDateModule } from '@angular/material/core';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';

import * as XLSX from 'xlsx';

@Component({
  selector: 'app-calendario',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatCardModule,
    MatInputModule,
    MatButtonModule,
    MatFormFieldModule,
    MatDatepickerModule,
    MatNativeDateModule,
    MatIconModule,
    MatProgressSpinnerModule
  ],
  templateUrl: './calendario.html',
  styleUrls: ['./calendario.css'],
  // Encapsulation None es REQUISITO para que los estilos personalizados afecten al calendario de Material
  encapsulation: ViewEncapsulation.None 
})
export class CalendarioComponent implements OnInit {
  nombreBusqueda: string = '';
  datosExcel: any[] = [];
  cargando: boolean = false;
  excelCargado: boolean = false;
  
  // Guardaremos los francos encontrados como strings en formato 'YYYY-MM-DD' para buscarlos rápido
  francosEmpleado: Set<string> = new Set<string>(); 
  busquedaRealizada: boolean = false;

  ngOnInit() {
    // Cargar el Excel automáticamente desde assets
    this.cargarExcelDesdeAssets();
  }

  // Cargar Excel desde los assets automáticamente
  private cargarExcelDesdeAssets() {
    this.cargando = true;
    fetch('assets/francos.xlsx')
      .then(response => response.arrayBuffer())
      .then(arrayBuffer => {
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        this.datosExcel = XLSX.utils.sheet_to_json(worksheet);
        this.excelCargado = true;
        this.cargando = false;
      })
      .catch(error => {
        console.error('Error cargando el Excel:', error);
        this.cargando = false;
      });
  }

  // 1. Leer el Excel y convertirlo a JSON (para carga manual si es necesario)
  onFileSelected(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    this.cargando = true;
    const fileReader = new FileReader();

    fileReader.onload = (e: any) => {
      const arrayBuffer = e.target.result;
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      this.datosExcel = XLSX.utils.sheet_to_json(worksheet);
      this.excelCargado = true;
      this.cargando = false;
    };

    fileReader.readAsArrayBuffer(file);
  }

  // 2. Buscar al empleado y mapear sus días Franco
  buscarFrancos() {
    if (!this.nombreBusqueda.trim()) return;
    
    this.francosEmpleado.clear();
    this.busquedaRealizada = true;
    const termino = this.nombreBusqueda.toLowerCase().trim();

    // Buscamos la fila del empleado
    const filaEmpleado = this.datosExcel.find(fila => 
      Object.values(fila).some(val => String(val).toLowerCase().includes(termino))
    );

    if (filaEmpleado) {
      // Recorremos las columnas de esa fila para buscar dónde dice "F" (Franco)
      Object.keys(filaEmpleado).forEach(columna => {
        const valorCelda = String(filaEmpleado[columna]).trim().toUpperCase();
        
        if (valorCelda === 'F' || valorCelda === 'FRANCO') {
          // Intentamos convertir el nombre de la columna en una fecha válida.
          // Esto asume que tus columnas del Excel se llaman como las fechas (ej: "2026-05-15" o "15/05/2026")
          const fechaParseada = new Date(columna);
          if (!isNaN(fechaParseada.getTime())) {
            this.francosEmpleado.add(this.formatDate(fechaParseada));
          }
        }
      });
    }
  }

  // Función ayudante para estandarizar las fechas a 'YYYY-MM-DD'
  private formatDate(date: Date): string {
    const d = new Date(date);
    let month = '' + (d.getMonth() + 1);
    let day = '' + d.getDate();
    const year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
  }

  // 3. Función inyectada en el <mat-calendar> de Angular Material
  // Se ejecuta para CADA día del mes. Si el día está en nuestro Set de francos, le aplica la clase CSS.
  dateClass: MatCalendarCellClassFunction<Date> = (cellDate, view) => {
    if (view === 'month') {
      const dateStr = this.formatDate(cellDate);
      return this.francosEmpleado.has(dateStr) ? 'dia-franco-custom' : '';
    }
    return '';
  };
}