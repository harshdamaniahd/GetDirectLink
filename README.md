## get-direct-link

![alt text](https://sharepointhd.files.wordpress.com/2019/03/ezgif.com-video-to-gif-2.gif)
Sample SharePoint Framework (SPFx) solution which gives the end-user the ability to just get a copy link with ID that doesnt break item level permission and it also generates url for pdf with id. This is helpful when document is renamed and moved in same site collection, thus not breaking page linking.This is done using a CommandSet.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO


