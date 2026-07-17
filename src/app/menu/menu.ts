import { Component, input, output } from '@angular/core';
import { RouterLink, RouterLinkActive } from '@angular/router';

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
      click: "fisico",
      icon: "fa fa-person-skating",
      label: "Físico",
    }
  ];

  onItemClick(item: MenuItem) {
    if (item.click) {
      console.log('Abrir modal:', item.click);
      // Aqui você poderá adicionar a lógica para abrir seu modal
    }
  }

  toggleCollaps(): void {
    this.changeMenu.emit(!this.ismenu());
  }

  closeSidenav(): void {
    this.changeMenu.emit(true);
  }
}
