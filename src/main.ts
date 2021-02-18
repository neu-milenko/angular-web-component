import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app/app.module';

import { environment } from './environments/environment';

if (environment.production) {
  enableProdMode();
}

function getGlobal(): any {
  return typeof self !== 'undefined'
    ? self
    : typeof window !== 'undefined'
    ? window
    // : typeof global !== "undefined"
    // ? global
    : undefined;
}

const g = getGlobal();

Office.initialize = () => {

  g.dialogCallback = (asyncResult : Office.AsyncResult<Office.Dialog>, event: any ): void => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      return;
    }

    const dialog = asyncResult.value;

    /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      console.log(`closing dialog ${arg}`)

      // dialog.close();
      event.completed();
    });
  };

  g.openDialog =(event) => {

    console.log(`opening dialog ${event}`)

    const h = Math.trunc((640 * 100) / screen.height);
    const w = Math.trunc((1296 * 100) / screen.width);

    Office.context.ui.displayDialogAsync(environment.ADD_IN_HOST,
      { height: h, width: w, displayInIframe: true },
        (result) => { g.dialogCallback(result, event); }
      );
    };

  console.log(`app bootstrap`);

  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(err => console.error(err));
};

Office.onReady((info) => {
  console.log(`onReady ${info}`);
  console.dir(info);

  window.onclose = () => {
    console.log('onclose');
    Office.context.ui.messageParent(true);
  };
});
