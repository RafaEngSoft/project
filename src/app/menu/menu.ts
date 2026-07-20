import { Component, input, output } from '@angular/core';
import { RouterLink, RouterLinkActive } from '@angular/router';
import { ModalService } from '../modal.service';

export interface MenuItem {
  routerLink?: string;
  click?: string;
  icon: string;
  label: string;
}

@Component({
  selector: 'app-menu',
  imports: [RouterLink, RouterLinkActive],
  templateUrl: './menu.html',
  styleUrl: './menu.css',
})
export class Menu {
  ismenu = input.required<boolean>();
  changeMenu = output<boolean>();

  items: MenuItem[] = [
    {
      routerLink: "fisico",
      icon: "fa fa-person-skating",
      label: "Físico",
    },
    {
      click: "curvaS",
      icon: "fa fa-chart-line",
      label: "Curva S",
    }
  ];

  constructor(private modalService: ModalService) {}

  onItemClick(item: MenuItem) {
    if (item.click === 'curvaS') {
      console.log('Abrir modal:', item.click);
      this.modalService.showSCurveModal.set(true);
    }
  }

  toggleCollaps(): void {
    this.changeMenu.emit(!this.ismenu());
  }

  closeSidenav(): void {
    this.changeMenu.emit(true);
  }
}
