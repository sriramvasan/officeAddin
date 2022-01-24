import { HashLocationStrategy, LocationStrategy } from "@angular/common";
import { NgModule } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import {AppComponent} from "./app.component";
import { HeaderComponent } from './header/header.component';
import { TransformComponent } from './transform/transform.component';
import { HomeComponent } from './home/home.component';
import { AppRoutingModule } from "./app-routing.module";
import { DefaultComponent } from './default/default.component';
import { ImportComponent } from './import/import.component';
import { FlattenComponent } from './flatten/flatten.component';
import { FormsModule } from "@angular/forms";
import { MultiplyComponent } from './multiply/multiply.component';
import { AlertComponent } from './shared/alert/alert.component';


@NgModule({
  declarations: [AppComponent, 
    TransformComponent,
    HomeComponent,
    DefaultComponent,
  HeaderComponent,
  ImportComponent,
  FlattenComponent,
  MultiplyComponent,
  AlertComponent],
  imports: [BrowserModule,
  AppRoutingModule,
FormsModule],
  bootstrap: [AppComponent],
  providers :[ {
     provide: LocationStrategy, useClass: HashLocationStrategy },
    ]
})
export class AppModule {}