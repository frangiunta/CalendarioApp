import { Component, OnInit, ViewEncapsulation, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

import { MatCardModule } from '@angular/material/card';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatCalendarCellClassFunction, MatDatepickerModule } from '@angular/material/datepicker';
import { MatNativeDateModule } from '@angular/material/core';
import { MatIconModule } from '@angular/material/icon';
import { MatListModule } from '@angular/material/list';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { MatAutocompleteModule } from '@angular/material/autocomplete';

import * as XLSX from 'xlsx';

import francosData from '../../assets/grilla_completa_2025_2027.json';
import backofficeData from '../../assets/gemini-code-1780341396539.json';

interface FilaExcel {
  Nombre?: string;
  empleado?: string;
  fecha?: string;
  trabaja?: boolean;
  estado?: string;
  horarios?: string[];
  [key: string]: any;
}

interface RawBackofficeCronograma {
  cronograma: Record<string, Array<{
    fecha: string;
    dia: string;
    turnos: Array<{ nombre: string; horario: string }>;
  }>>;
}

interface BackofficeRow {
  fecha: string;
  dia: string;
  nombre: string;
  horario: string;
  imagenUrl?: string;
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
    ,MatListModule
  ],
  templateUrl: './calendario.html',
  styleUrls: ['./calendario.css'],

  encapsulation: ViewEncapsulation.None 
})
export class CalendarioComponent implements OnInit {
  nombreBusqueda: string = '';
  datosExcel: FilaExcel[] = [];
  cargando: boolean = false;
  excelCargado: boolean = false;
  nombresEmpleados: string[] = [];
  opcionesFiltradas: string[] = [];

  backofficeRows: BackofficeRow[] = [];
  backofficeFilter: string = '';
  backofficeRowsFiltrados: BackofficeRow[] = [];
  backofficeMaxRows: number = 250;
  backofficeCollapsed: boolean = false;

  francosEmpleado: Set<string> = new Set<string>(); 
  busquedaRealizada: boolean = false;

  constructor(private cdr: ChangeDetectorRef) {}

  ngOnInit() {
    this.cargarDatosDesdeAssets();
    this.cargarBackofficeDesdeAssets();
  }

  toggleBackoffice() {
    this.backofficeCollapsed = !this.backofficeCollapsed;
  }

  private getImagenUrlPorNombre(nombreCompleto: string): string {
    if (!nombreCompleto) {
      return './assets/empleados/default.png'; // Imagen por defecto
    }
    // Extrae el nombre, por ej. "DEBORA" de "1 DEBORA"
    const nombreRaw = this.extraerNombreEmpleado(nombreCompleto).split(' ')[0];
    const nombre = this.normalizarTexto(nombreRaw); // Quita tildes y pasa a minúsculas
    if (!nombre) {
        return './assets/empleados/default.png';
    }
    return `./assets/empleados/${nombre}.jpg`;
  }

  private extraerNombreEmpleado(empRaw: string): string {
    return String(empRaw || '').trim().replace(/^\s*\d+\s*/,'').replace(/\*/g,'').trim();
  }

  seleccionarEmpleado(nombre: string) {
    this.nombreBusqueda = nombre;
    this.filtrarOpciones(nombre);
    this.buscarFrancos();
  }

  private cargarDatosDesdeAssets() {
    this.cargando = true;
    
    try {

      const data: FilaExcel[] = (francosData as any).default ? (francosData as any).default : francosData;

      this.datosExcel = data;

      if (data.length > 0 && (data as any)[0].empleado !== undefined) {
        const nombresSet = new Set<string>();
        for (const fila of data as any[]) {
          const nombre = this.extraerNombreEmpleado(String(fila.empleado || ''));
          if (nombre) nombresSet.add(nombre);
        }
        this.nombresEmpleados = Array.from(nombresSet).sort();
      } else {

        this.nombresEmpleados = data
          .map((fila: FilaExcel) => String(fila.Nombre || ''))
          .filter((nombre: string) => nombre.trim() !== '');
      }

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

      if (this.datosExcel.length > 0 && (this.datosExcel as any)[0].empleado !== undefined) {
        const nombresSet = new Set<string>();
        for (const fila of this.datosExcel as any[]) {
          const nombre = this.extraerNombreEmpleado(String(fila.empleado || ''));
          if (nombre) nombresSet.add(nombre);
        }
        this.nombresEmpleados = Array.from(nombresSet).sort();
      } else {
        this.nombresEmpleados = this.datosExcel
          .map(fila => String(fila.Nombre || ''))
          .filter(nombre => nombre.trim() !== '');
      }
      this.opcionesFiltradas = this.nombresEmpleados.slice(0, 50);
      
      this.excelCargado = true;
      this.cargando = false;
    };

    fileReader.readAsArrayBuffer(file);
  }

  private cargarBackofficeDesdeAssets() {
    try {
      const raw: RawBackofficeCronograma = (backofficeData as any).default ? (backofficeData as any).default : backofficeData;
      this.backofficeRows = Object.values(raw.cronograma)
        .flatMap((dias) =>
          dias.flatMap((dia) =>
            (Array.isArray(dia.turnos) ? dia.turnos : []).map((turno) => {
              const nombreCompleto = String(turno.nombre || '').trim();
              return {
                fecha: dia.fecha,
                dia: dia.dia,
                nombre: nombreCompleto,
                horario: String(turno.horario || '').trim(),
                imagenUrl: this.getImagenUrlPorNombre(nombreCompleto)
              };
            })
          )
        );
      this.actualizarBackofficeFiltrados();
    } catch (error) {
      console.error('Error procesando el JSON de backoffice:', error);
    }
  }

  actualizarBackofficeFiltrados() {
    const filtro = this.normalizarTexto(this.backofficeFilter || '');
    this.backofficeRowsFiltrados = this.backofficeRows
      .filter((row) => {
        if (!filtro) return true;
        const texto = [row.fecha, row.dia, row.nombre, row.horario].join(' ');
        return this.normalizarTexto(texto).includes(filtro);
      })
      .slice(0, filtro ? this.backofficeRows.length : this.backofficeMaxRows);
  }

  formatHorarios(horarios: string[] = []): string {
    return horarios.join(', ');
  }

  private normalizarTexto(texto: string): string {
    return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  }

  filtrarOpciones(valor: string) {
    const valorNormalizado = this.normalizarTexto(valor);
    this.opcionesFiltradas = this.nombresEmpleados
      .filter(nombre => this.normalizarTexto(nombre).includes(valorNormalizado))
      .slice(0, 50); 
  }

  limpiarBusqueda() {
    this.nombreBusqueda = '';
    this.opcionesFiltradas = this.nombresEmpleados.slice(0, 50);
    this.francosEmpleado.clear();
    this.busquedaRealizada = false;
  }

  buscarFrancos() {
    if (!this.nombreBusqueda.trim()) return;
    
    this.francosEmpleado.clear();
    this.busquedaRealizada = true;
    const termino = this.normalizarTexto(this.nombreBusqueda);
    const exactMatch = this.nombresEmpleados.find(emp => this.normalizarTexto(emp) === termino);

    if (this.datosExcel.length > 0 && (this.datosExcel as any)[0].empleado !== undefined) {
      for (const fila of this.datosExcel as any[]) {
        const nombre = this.extraerNombreEmpleado(String(fila.empleado || ''));
        if (!nombre) continue;

        const nombreNormalizado = this.normalizarTexto(nombre);
        const coincidencia = exactMatch ? nombreNormalizado === termino : nombreNormalizado.includes(termino);

        if (coincidencia) {
          const trabaja = fila.trabaja;

          const esFranco = trabaja === false || String(trabaja).toLowerCase() === 'false';

          if (esFranco) {
            const fecha = String(fila.fecha || '').trim();
            if (/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
              this.francosEmpleado.add(fecha);
            } else {
              const fechaParseada = new Date(fecha);
              if (!isNaN(fechaParseada.getTime())) {
                this.francosEmpleado.add(this.formatDate(fechaParseada));
              }
            }
          }
        }
      }
    } else {

      const filaEmpleado = this.datosExcel.find((fila: FilaExcel) => 
        this.normalizarTexto(String(fila.Nombre || '')).includes(termino)
      );

      if (filaEmpleado) {

        Object.keys(filaEmpleado).forEach(columna => {
          const valorCelda = String(filaEmpleado[columna]).trim().toUpperCase();
          
          if (valorCelda === 'F' || valorCelda === 'FRANCO') {

            if (/^\d{4}-\d{2}-\d{2}$/.test(columna)) {
              this.francosEmpleado.add(columna);
            } else {

              const fechaParseada = new Date(columna);
              
              if (!isNaN(fechaParseada.getTime())) {

                const fechaUTC = new Date(fechaParseada.getTime() + fechaParseada.getTimezoneOffset() * 60000);
                this.francosEmpleado.add(this.formatDate(fechaUTC));
              }
            }
          }
        });
      }
    }
  }

  private formatDate(date: Date): string {
    const d = new Date(date);
    let month = '' + (d.getMonth() + 1);
    let day = '' + d.getDate();
    const year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
  }

  private formatIcsDate(date: Date): string {
    const d = new Date(date);
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${year}${month}${day}`;
  }

  private formatIcsDateTime(date: Date): string {
    const d = new Date(date);
    const year = d.getUTCFullYear();
    const month = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    const hours = String(d.getUTCHours()).padStart(2, '0');
    const minutes = String(d.getUTCMinutes()).padStart(2, '0');
    const seconds = String(d.getUTCSeconds()).padStart(2, '0');
    return `${year}${month}${day}T${hours}${minutes}${seconds}Z`;
  }

  exportCalendarioIcs() {
    if (!this.nombreBusqueda || this.francosEmpleado.size === 0) return;

    const ics = this.buildIcs();
    const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${this.nombreBusqueda.replace(/\s+/g, '_') || 'francos'}.ics`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  private buildIcs(): string {
    const now = new Date();
    const dtstamp = this.formatIcsDateTime(now);
    const summary = `Franco - ${this.nombreBusqueda}`;
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-,,CalendarioFrancoApp',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH'
    ];

    const fechas = Array.from(this.francosEmpleado).sort();
    fechas.forEach((fecha, index) => {
      const dtstart = fecha.replace(/-/g, '');
      const nextDay = new Date(`${fecha}T00:00:00`);
      nextDay.setDate(nextDay.getDate() + 1);
      const dtend = this.formatIcsDate(nextDay);

      lines.push('BEGIN:VEVENT');
      lines.push(`UID:franco-${index}-${dtstamp}@calendario`);
      lines.push(`DTSTAMP:${dtstamp}`);
      lines.push(`SUMMARY:${summary}`);
      lines.push(`DESCRIPTION:Franco`);
      lines.push(`DTSTART;VALUE=DATE:${dtstart}`);
      lines.push(`DTEND;VALUE=DATE:${dtend}`);
      lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }


  dateClass: MatCalendarCellClassFunction<Date> = (cellDate, view) => {
    if (view === 'month') {
      const dateStr = this.formatDate(cellDate);
      return this.francosEmpleado.has(dateStr) ? 'dia-franco-custom' : '';
    }
    return '';
  };
}