import { Injectable } from '@angular/core';
import { Subject } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class MessagesService {

  private isTablePresent :boolean;
  tableStatusChanged  = new Subject<boolean>();
  messageChanged = new Subject<string>();
  isMessageSent = new Subject<boolean>();

  messages: string[]=["default Page"];

  constructor() { }

  getMessage(){
    return this.messages.slice();
  }

  setMessage(msg:string){
    this.messages.push(msg); 
    this.messageChanged.next(msg);
  }

  setTableStatus(status:boolean){
    this.isTablePresent = status;
    this.tableStatusChanged.next(this.isTablePresent);
  }
  
  getTableStatus(){
    return this.isTablePresent;
  }

}
