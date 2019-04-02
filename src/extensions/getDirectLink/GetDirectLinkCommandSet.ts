import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import GetDirectLink from './components/GetDirectLink';
import { sp, SharingInformation } from "@pnp/sp";
import * as strings from 'GetDirectLinkCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGetDirectLinkCommandSetProperties {

}

const LOG_SOURCE: string = 'GetDirectLinkCommandSet';

export default class GetDirectLinkCommandSet extends BaseListViewCommandSet<IGetDirectLinkCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized GetDirectLinkCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    debugger;
    const compareOneCommand: Command = this.tryGetCommand('Command_DirectLink');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
      compareOneCommand.title = strings.Command_DirectLink;

    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'Command_DirectLink':
        //Gets Full Site Url      
        let relativeUrl = event.selectedRows[0].getValueByName('FileRef');
        //Gets FileName
        let fileName = event.selectedRows[0].getValueByName('FileLeafRef');
        //Gets FileExtension
        let fileExtension = fileName.split(".")[1];
        //Gets the filename to display on dialog
        let linkTo=fileName.length > 20 ? fileName.substring(0,8)+"..." + fileName.substring(fileName.length-8, fileName.length):fileName;
        let url;
        //Returns file url with id
        if (fileExtension !== "pdf") {
          url = await sp.web.getFolderByServerRelativeUrl(relativeUrl)
            .getSharingInformation().then((result: SharingInformation) => {
              return result.directUrl
            }).catch(e => {
              return e
            })
        }
        else {
          url = await sp.web.getFileByServerRelativeUrl(relativeUrl).listItemAllFields.select("ServerRedirectedEmbedUrl").get().then((items => {
            return (items.ServerRedirectedEmbedUrl);
          }))
        }
        let msg = url["length"] > 255 ? strings.UrlMsg : "";
        const callout: GetDirectLink = new GetDirectLink();
        callout.fileName = linkTo;
        callout.absolutePath = url;
        callout.msg = msg;
        callout.fileNameToolTip=fileName;
        callout.show();
        console.log("show")
        break;
        default:
        throw new Error('Unknown command');
    }
  }
}
