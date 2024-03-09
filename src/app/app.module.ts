import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { AppComponent } from './app.component';
// Import SpreadJS module
import { SpreadSheetsModule } from "@mescius/spread-sheets-angular";

@NgModule({
  declarations: [
    AppComponent
  ],
  // Import SpreadJS module
  imports: [
    BrowserModule, SpreadSheetsModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
