import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
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
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1':
        // let siteUrl = this.context.pageContext.site.absoluteUrl;
        // let endIndex = siteUrl.lastIndexOf('/sites/');
        // let rootSiteUrl = siteUrl.substring(0, endIndex);
        let relativeUrl = event.selectedRows[0].getValueByName('FileRef');
        let fileName = event.selectedRows[0].getValueByName('FileLeafRef');
        let fileExtension = fileName.split(".")[1];
        let linkTo=fileName.length > 20 ? fileName.substring(0,8)+"..." + 
        fileName.substring(fileName.length-8, fileName.length)    :
        fileName;

        let url;
        if (fileExtension !== "pdf") {
          url = await sp.web.getFolderByServerRelativeUrl(relativeUrl)
            .getSharingInformation().then((result: SharingInformation) => {
              return result.directUrl
            }).catch(e => {
              return e
            })
        }
        else {
          url = await sp.web.getFileByServerRelativeUrl(relativeUrl).listItemAllFields.get().then((items => {
            return (items.ServerRedirectedEmbedUrl);
          }))
        }
        //let msg = url.hasOwnProperty("isHttpRequestError") ? url["message"] : ""
        let msg = url["length"] > 255 ? "Url may not contain more than 255 chars" : "";
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
