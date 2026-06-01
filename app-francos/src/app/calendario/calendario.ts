import { Component, OnInit, ViewEncapsulation, ChangeDetectorRef } from '@angular/core';
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

// Importamos el JSON directamente (Angular lo empaquetará, evitando cualquier error 404 de red)
// @ts-ignore
import francosData from '../../assets/francos.json';

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

  constructor(private cdr: ChangeDetectorRef) {}

  ngOnInit() {
    // Cargar los datos automáticamente desde assets
    this.cargarDatosDesdeAssets();
  }

  // Cargar JSON desde los assets automáticamente (mucho más rápido)
  private cargarDatosDesdeAssets() {
    this.cargando = true;
    
    try {
      // Dependiendo del compilador, los datos pueden venir en .default o directamente en el objeto
      const data: FilaExcel[] = (francosData as any).default ? (francosData as any).default : francosData;

      this.datosExcel = data;
      
      // Extraemos solo los nombres, descartamos vacíos y valores que sean estrictamente numéricos (ej. "2", "4")
      this.nombresEmpleados = data
        .map((fila: FilaExcel) => String(fila.Nombre || ''))
        .filter((nombre: string) => nombre.trim() !== '' && isNaN(Number(nombre)));
        
      // Limitamos a 50 opciones para evitar que el navegador se congele renderizando HTML masivo
      this.opcionesFiltradas = this.nombresEmpleados.slice(0, 50);
      
      this.excelCargado = true;
      this.cargando = false;
      this.cdr.detectChanges();
    } catch (error) {
      console.error('Error procesando los datos JSON importados:', error);
      this.cargando = false;
      this.cdr.detectChanges();
    }
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
      
      // Actualizamos la lista también en caso de carga manual (ignorando números)
      this.nombresEmpleados = this.datosExcel
        .map(fila => String(fila.Nombre || ''))
        .filter(nombre => nombre.trim() !== '' && isNaN(Number(nombre)));
      this.opcionesFiltradas = this.nombresEmpleados.slice(0, 50);
      
      this.excelCargado = true;
      this.cargando = false;
    };

    fileReader.readAsArrayBuffer(file);
  }

  // Elimina acentos/tildes para que buscar "gaston" encuentre "GASTÓN"
  private normalizarTexto(texto: string): string {
    return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  }

  // Filtra las opciones del autocompletado basándose en lo que escribe el usuario
  filtrarOpciones(valor: string) {
    const valorNormalizado = this.normalizarTexto(valor);
    this.opcionesFiltradas = this.nombresEmpleados
      .filter(nombre => this.normalizarTexto(nombre).includes(valorNormalizado))
      .slice(0, 50); // Mantener un máximo de 50 resultados garantiza fluidez
  }

  // Limpia el buscador y el calendario
  limpiarBusqueda() {
    this.nombreBusqueda = '';
    this.opcionesFiltradas = this.nombresEmpleados.slice(0, 50);
    this.francosEmpleado.clear();
    this.busquedaRealizada = false;
  }

  // 2. Buscar al empleado y mapear sus días Franco
  buscarFrancos() {
    if (!this.nombreBusqueda.trim()) return;
    
    this.francosEmpleado.clear();
    this.busquedaRealizada = true;
    const termino = this.normalizarTexto(this.nombreBusqueda);

    // Buscamos la fila del empleado (Consultando únicamente la propiedad "Nombre", mucho más rápido)
    const filaEmpleado = this.datosExcel.find((fila: FilaExcel) => 
      this.normalizarTexto(String(fila.Nombre || '')).includes(termino)
    );

    if (filaEmpleado) {
      // Recorremos las columnas de esa fila para buscar dónde dice "F" (Franco)
      Object.keys(filaEmpleado).forEach(columna => {
        const valorCelda = String(filaEmpleado[columna]).trim().toUpperCase();
        
        if (valorCelda === 'F' || valorCelda === 'FRANCO') {
          // Si la fecha ya viene en formato YYYY-MM-DD directo del JSON, la agregamos al instante
          if (/^\d{4}-\d{2}-\d{2}$/.test(columna)) {
            this.francosEmpleado.add(columna);
          } else {
            // Si viene de un Excel manual, intentamos parsear el nombre de la columna como fecha
            const fechaParseada = new Date(columna);
            
            if (!isNaN(fechaParseada.getTime())) {
              // Utilizamos UTC para evitar que el ajuste de zona horaria local retroceda 1 día la fecha
              const fechaUTC = new Date(fechaParseada.getTime() + fechaParseada.getTimezoneOffset() * 60000);
              this.francosEmpleado.add(this.formatDate(fechaUTC));
            }
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