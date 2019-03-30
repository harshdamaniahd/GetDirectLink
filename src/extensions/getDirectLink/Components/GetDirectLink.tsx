import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './GetDirectLinkCopy.module.scss'
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { getId } from 'office-ui-fabric-react/lib/Utilities';


interface IGetDirectLinkProps {
    fileName: string;
    absolutePath: string;
    domElement: any;
    msg:any;
    fileNameToolTip:string;
    onDismiss: () => void;
}
interface IGetDirectLinkState{}


export default class GetDirectLinkComponent extends BaseDialog {
    public fileName: string;
    public absolutePath: any;
    public msg:any;
    public fileNameToolTip:string;
    public render(): void {        
      ReactDOM.render(<GetDirectLinkContent
        fileName={ this.fileName }
        absolutePath={ this.absolutePath }
        domElement={ document.activeElement.parentElement }
        onDismiss={this.onDismiss.bind(this)}
        msg={this.msg}
        fileNameToolTip={this.fileNameToolTip}
      />, this.domElement);
    }

    private onDismiss()
    {
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}

class GetDirectLinkContent extends 
  React.Component<IGetDirectLinkProps, IGetDirectLinkState> {
    private _hostId: string = getId('tooltipHost');

    constructor(props : IGetDirectLinkProps) {
      super(props);

      this.state = {
      };
    }

    public render(): JSX.Element {
      return (
          <div>
            <Callout   
                className="ms-CalloutExample-callout"
                ariaLabelledBy={'callout-label-1'}
                ariaDescribedBy={'callout-description-1'}
                role={'alertdialog'}                
                gapSpace={0}
                target={this.props.domElement}
                hidden={false}
                calloutWidth={320}
                preventDismissOnScroll={true}
                setInitialFocus={true}                
                onDismiss={this.onDismiss.bind(this)}>
                <div className={styles.justALinkContentContainer}>
                    <div className={styles.iconContainer} ><Icon iconName="CheckMark" className={styles.icon} /></div>
                    <TooltipHost content={this.props.fileNameToolTip} id={this._hostId} calloutProps={{ gapSpace: 0 }}>
          <div aria-labelledby={this._hostId} className={styles.fileName}>Link to '{this.props.fileName}' copied</div>
        </TooltipHost>
                   
                    <div className={styles.shareContainer}>
                        <TextField className={styles.filePathTextBox} value={this.props.absolutePath} />
                        <PrimaryButton text="Copy" onClick={this.btnCopyCLicked.bind(this)}
                        />
                    </div>
                    <Label className="ms-fontColor-red">{this.props.msg}</Label>                    
                </div>
            </Callout>
          </div>
      );
    }

    private onDismiss(ev: any)
    {
        this.props.onDismiss();
    }

    private btnCopyCLicked(): void {
        var el = document.createElement('textarea');
        el.value = this.props.absolutePath;
        el.setAttribute('readonly', '');
        el.style.position = 'absolute';
        el.style.left = '-9999px';
        document.body.appendChild(el);
        el.select();

        document.execCommand('copy');
        document.body.removeChild(el);
    }
}