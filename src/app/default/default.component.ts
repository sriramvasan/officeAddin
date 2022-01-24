import { Component,OnDestroy,OnInit } from '@angular/core';
import { Subscription } from 'rxjs';
import { MessagesService } from '../messages.service';

@Component({
  selector: 'app-default',
  templateUrl: './default.component.html',
  styleUrls: ['./default.component.css']
})
export class DefaultComponent implements OnInit, OnDestroy {
  Message : String  = this.messageService.getMessage().pop();
  subscription :Subscription ;
  
  constructor(private messageService:MessagesService) { }

  ngOnInit(): void {
    //  this.subscription = 
     this.messageService.messageChanged.subscribe(
        (msg : string)=>{
          this.Message = msg;
        }
      )    
      this.Message = this.messageService.getMessage().pop();
  }
  ngOnDestroy(){
    // this.subscription.unsubscribe();
  }

}
