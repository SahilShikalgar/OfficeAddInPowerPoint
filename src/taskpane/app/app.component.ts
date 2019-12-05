import { Component, OnInit } from "@angular/core";
import { async } from "q";
const template = require("./app.component.html");
/* global console, Office, require */

@Component({
  selector: "app-home",
  template
})
export default class AppComponent implements OnInit {
  welcomeMessage = "Welcome";
  image;

  ngOnInit() {
    if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
      console.log(Office);
      // Office.context.(Office.FileType.Text,null, (data) => {
      //   console.log("call back", data);
      // });
    }
    else {
      // Provide alternate flow/logic.
      console.log("not supported");
    }
  }

  async run() {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }

  async createPresentation() {
    PowerPoint.createPresentation();
  }

  onSelectFile($event) {
    this.readThis($event.target);
  }

  async insertImage() {
    const startIndex = this.image.indexOf("base64,");
    const copyBase64 = this.image.substr(startIndex + 7);
  
    await Office.context.document.setSelectedDataAsync(
      copyBase64,
      {
        coercionType: Office.CoercionType.Image
      },
      async result => {
        console.log(result);
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result);
        }
      }
    );
  }

  readThis(inputValue: any): void {
    var file:File = inputValue.files[0];
    var myReader:FileReader = new FileReader();
  
    myReader.onloadend = (e) => {
      this.image = myReader.result;
    }
    myReader.readAsDataURL(file);
  }

  insertTable() {
    
  }

  async getSlide() {

  // Can get content except images
    // await Office.context.document.getSelectedDataAsync(
    //   Office.CoercionType.SlideRange,
    //   null,
    //   result => {
    //     console.log(result);
    //     if (result.status === Office.AsyncResultStatus.Failed) {
    //       console.error(result.error.message);
    //     }
    //   }
    // );


  //get selected Text or Slide from powerpoint
    // Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, null, result => {
    //   console.log(result);
    // });

    Office.context.document.getSelectedDataAsync(Office.CoercionType.Image, result => {
      console.log(result);
    });
  }

  goToFirstSlide() {
    Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
        function (asyncResult) {
          console.log(asyncResult);
        });
  }

  goToLastSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
          function (asyncResult) {
            console.log(asyncResult);
          });
  }

  goToPreviousSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
          function (asyncResult) {
            console.log(asyncResult);
          });
  }

  goToNextSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
          function (asyncResult) {
            console.log(asyncResult);
          });
  }
}
