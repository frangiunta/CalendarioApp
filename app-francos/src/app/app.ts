import { Component } from '@angular/core';
import { CalendarioComponent } from './calendario/calendario';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CalendarioComponent],
  template: '<app-calendario></app-calendario>',
  styleUrl: './app.css'
})
export class App {}
