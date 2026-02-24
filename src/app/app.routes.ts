import { Routes } from '@angular/router';
import { MigrationPageComponent } from './pages/migration-page.component';
import { IntegrateSpPageComponent } from './pages/integrate-sp-page.component';

export const routes: Routes = [
  { path: '', pathMatch: 'full', redirectTo: 'migration' },
  { path: 'migration', component: MigrationPageComponent },
  { path: 'integrate-sp', component: IntegrateSpPageComponent },
  { path: '**', redirectTo: 'migration' }
];
