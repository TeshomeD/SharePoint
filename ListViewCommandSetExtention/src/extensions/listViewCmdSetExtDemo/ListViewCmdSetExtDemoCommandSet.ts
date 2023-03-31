import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ListViewCmdSetExtDemoCommandSetStrings';
import * as pnp from 'sp-pnp-js'

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListViewCmdSetExtDemoCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ListViewCmdSetExtDemoCommandSet';

export default class ListViewCmdSetExtDemoCommandSet extends BaseListViewCommandSet<IListViewCmdSetExtDemoCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ListViewCmdSetExtDemoCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareTwoCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareTwoCommand.visible = event.selectedRows.length > 1;
    }
    const compareThreeCommand: Command = this.tryGetCommand('COMMAND_3');
    if (compareThreeCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareThreeCommand.visible = event.selectedRows.length > 1;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        //Dialog.alert(`${this.properties.sampleTextOne}`);
       let title: string = event.selectedRows[0].getValueByName('Title');
       let status: string = event.selectedRows[0].getValueByName('Status'); 
       Dialog.alert(`Projecrt Name: ${title} - Current Status: ${status}% done`)
        break;
      case 'COMMAND_2':
        Dialog.alert(`${this.properties.sampleTextTwo}`);
        break;
      case 'COMMAND_3':
        Dialog.prompt(`Project Status Remarks`).then((value: string)=>{
          this.UpdateRemarks(event.selectedRows, value);
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  private UpdateRemarks(items: any, value:string):void{
    let batch = pnp.sp.createBatch();
    items.forEach(item => {
      pnp.sp.web.lists.getByTitle('ProjectsStatus').items
          .getById(item.getValueByName("ID")).inBatch(batch)
          .update({Remarks: value}).then(res=>{

          });
          batch.execute().then(res=> {
            location.reload();
          });
    });
  }
}
