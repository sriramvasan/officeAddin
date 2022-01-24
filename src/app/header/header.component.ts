import { Component, OnInit } from '@angular/core';
import { MessagesService } from '../messages.service';

@Component({
  selector: 'header',
  templateUrl: './header.component.html',
  styleUrls: ['./header.component.css']
})
export class HeaderComponent implements OnInit {

  isShown:boolean = false;
  messageDisplay : boolean = false;
  constructor(private messageService:MessagesService) {

   }

  ngOnInit(): void {
    
  }

  
}
