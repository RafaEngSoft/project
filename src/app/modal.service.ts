import { Injectable, signal } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class ModalService {
  showSCurveModal = signal<boolean>(false);
}
