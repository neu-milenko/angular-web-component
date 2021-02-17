import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app/app.module';

import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    // : typeof global !== "undefined"
    // ? global
    : undefined;
}

const g = getGlobal() as any;

Office.initialize = () => {

  g.dialogCallback = function dialogCallback(asyncResult : Office.AsyncResult<Office.Dialog>, event: any ) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed)
      return;

    let dialog = asyncResult.value;

    /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
    dialog.addEventHandler(Office.EventType.DialogEventReceived, function(arg) {
      console.log(`closing dialog ${arg}`)

      //dialog.close();
      event.completed();
    });
  };

  g.openDialog = function openDialog(event) {

    console.log(`opening dialog ${event}`)

    let h = Math.trunc((640 * 100) / screen.height);
    let w = Math.trunc((1296 * 100) / screen.width);

    Office.context.ui.displayDialogAsync(environment.ADD_IN_HOST,
      { height: h, width: w, displayInIframe: true },
      function(result) { g.dialogCallback(result, event) }
      );
    };

  console.log(`app bootstrap`);

  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(err => console.error(err));
};

// Office.onReady((info) => {
//   console.log(`onReady ${info}`);

//   window.onclose = () => {
//     console.log('onclose');
//     Office.context.ui.messageParent(true);
//   };
// });
