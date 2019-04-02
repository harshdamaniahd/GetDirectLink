declare interface IGetDirectLinkCommandSetStrings {
  Command_DirectLink: string;
  Copy:string;
  UrlMsg:string;
  LinkTo:string;
  Copied:string;

}

declare module 'GetDirectLinkCommandSetStrings' {
  const strings: IGetDirectLinkCommandSetStrings;
  export = strings;
}
