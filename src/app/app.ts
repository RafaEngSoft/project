import { Component, HostListener, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { Menu } from "./menu/menu";

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, Menu],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly title = signal('project');
  ismenu = signal<boolean>(false);
  screenWidth = signal<number>(typeof window !== 'undefined' ? window.innerWidth : 1200);

  @HostListener("window:resize")
  onResize() {
    if (typeof window !== 'undefined') {
      this.screenWidth.set(window.innerWidth);
      if (this.screenWidth() < 768) {
        this.ismenu.set(true);
      }
    }
  }

  changeMenu(ismenu: boolean): void {
     this.ismenu.set(ismenu);
  }
}
