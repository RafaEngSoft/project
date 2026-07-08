import { Routes } from '@angular/router';
import { homeComponent } from './home/home';

export const routes: Routes = [
    // Se o usuário acessar o site sem digitar nada no final, manda para /home
    { path: '', redirectTo: 'home', pathMatch: 'full' },

    // Rota que carrega o seu componente Home
    { path: 'home', component: homeComponent }
];