import { Component, EventEmitter, Input, OnDestroy, OnInit, Output } from '@angular/core';

@Component({
  selector: 'app-alert',
  templateUrl: './alert.component.html',
  styleUrls: ['./alert.component.css']
})
export class AlertComponent implements OnInit, OnDestroy {

  constructor() { }

  ngOnInit(): void {
  }

    isVisible : boolean = false;

    @Input() message : string;
    @Output() close = new EventEmitter<void>();
    
    onClose(){
      this.message = null;
      this.close.emit();
    }
    
    ngOnDestroy(): void {
        this.close.unsubscribe();
    }

}
