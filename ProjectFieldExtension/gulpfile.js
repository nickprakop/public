'use strict';

const build = require('@microsoft/sp-build-web');
const path = require('path');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (!generatedConfiguration.resolve.alias) {
      generatedConfiguration.resolve.alias = {};
    }
    //root src folder   
    generatedConfiguration.resolve.alias['@src'] = path.resolve(__dirname, 'lib')
    return generatedConfiguration;
  }
});

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));



