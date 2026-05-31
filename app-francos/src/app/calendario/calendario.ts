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
import { MatAutocompleteModule } from '@angular/material/autocomplete';

import * as XLSX from 'xlsx';

// Interfaz para estructurar los datos extraídos del Excel
interface FilaExcel {
  Nombre?: string;
  [fecha: string]: string | number | undefined | Date;
}

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
    MatProgressSpinnerModule,
    MatAutocompleteModule
  ],
  templateUrl: './calendario.html',
  styleUrls: ['./calendario.css'],
  // Encapsulation None es REQUISITO para que los estilos personalizados afecten al calendario de Material
  encapsulation: ViewEncapsulation.None 
})
export class CalendarioComponent implements OnInit {
  nombreBusqueda: string = '';
  datosExcel: FilaExcel[] = [];
  cargando: boolean = false;
  excelCargado: boolean = false;
  nombresEmpleados: string[] = [];
  opcionesFiltradas: string[] = [];
  
  // Guardaremos los francos encontrados como strings en formato 'YYYY-MM-DD' para buscarlos rápido
  francosEmpleado: Set<string> = new Set<string>(); 
  busquedaRealizada: boolean = false;

  ngOnInit() {
    // Cargar los datos automáticamente desde assets
    this.cargarDatosDesdeAssets();
  }

  // Cargar JSON desde los assets automáticamente (mucho más rápido)
  private cargarDatosDesdeAssets() {
    this.cargando = true;
    fetch('assets/francos.json')
      .then(response => response.json())
      .then((data: FilaExcel[]) => {
        this.datosExcel = data;
        // Extraemos solo los nombres y descartamos valores vacíos
        this.nombresEmpleados = data.map(fila => String(fila.Nombre || '')).filter(nombre => nombre.trim() !== '');
        this.opcionesFiltradas = this.nombresEmpleados;
        
        this.excelCargado = true;
        this.cargando = false;
      })
      .catch(error => {
        console.error('Error cargando los datos:', error);
        this.cargando = false;
      });
  }

  // 1. Leer el Excel y convertirlo a JSON (para carga manual si es necesario)
  onFileSelected(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;

    this.cargando = true;
    const fileReader = new FileReader();

    fileReader.onload = (e: any) => {
      const arrayBuffer = e.target.result;
      const data = new Uint8Array(arrayBuffer);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      this.datosExcel = XLSX.utils.sheet_to_json<FilaExcel>(worksheet);
      // Actualizamos la lista también en caso de carga manual
      this.nombresEmpleados = this.datosExcel.map(fila => String(fila.Nombre || '')).filter(nombre => nombre.trim() !== '');
      this.opcionesFiltradas = this.nombresEmpleados;
      
      this.excelCargado = true;
      this.cargando = false;
    };

    fileReader.readAsArrayBuffer(file);
  }

  // Filtra las opciones del autocompletado basándose en lo que escribe el usuario
  filtrarOpciones(valor: string) {
    const valorMin = valor.toLowerCase().trim();
    this.opcionesFiltradas = this.nombresEmpleados.filter(nombre => nombre.toLowerCase().includes(valorMin));
  }

  // 2. Buscar al empleado y mapear sus días Franco
  buscarFrancos() {
    if (!this.nombreBusqueda.trim()) return;
    
    this.francosEmpleado.clear();
    this.busquedaRealizada = true;
    const termino = this.nombreBusqueda.toLowerCase().trim();

    // Buscamos la fila del empleado
    const filaEmpleado = this.datosExcel.find((fila: FilaExcel) => 
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
            // Utilizamos UTC para evitar que el ajuste de zona horaria local retroceda 1 día la fecha
            const fechaUTC = new Date(fechaParseada.getTime() + fechaParseada.getTimezoneOffset() * 60000);
            this.francosEmpleado.add(this.formatDate(fechaUTC));
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