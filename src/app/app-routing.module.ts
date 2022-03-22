import { NgModule } from '@angular/core';
import {  RouterModule, Routes } from '@angular/router';
import { DefaultComponent } from './default/default.component';
import { FlattenComponent } from './flatten/flatten.component';
import { HomeComponent } from './home/home.component';
import { ImportComponent } from './import/import.component';
import { MultiplyComponent } from './multiply/multiply.component';
import { TransformComponent } from './transform/transform.component';


const routes :Routes = [

  {path : 'home' , component : HomeComponent},
  {path : 'transform' , component : TransformComponent},
  {path: 'default', component: DefaultComponent},
  {path :'import',component:ImportComponent},
  {path :'flatten',component:FlattenComponent},
  {path: 'multiply',component:MultiplyComponent},
  {path: '' , redirectTo:'/home',pathMatch:'full'},
  {path : '**' , redirectTo:'/default', pathMatch:'full'}
]

@NgModule({
  declarations: [],
  imports: [
    RouterModule.forRoot(routes , {useHash :true})
  ],
  exports :[RouterModule]
})
export class AppRoutingModule { }
