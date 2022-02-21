const async = require('async');
const url = require('url');
const xbytes = require('xbytes');
const { sp } = require('@pnp/sp-commonjs');
const { PnpNode } = require('sp-pnp-node');

const fileTypes = {
  pdf: 'file-pdf',
  html: 'file-code',
  csv: 'file-csv',
  zip: 'file-archive',
  jpg: 'image',
  png: 'image',
  gif: 'image',
  xlsx: 'file-excel',
  docx: 'file-word',
  doc: 'file-word',
  ppt: 'file-powerpoint',
  pptx: 'file-powerpoint',
  html: 'file',
  json: 'file',
  log: 'file'
};

let optionsHash = '';
let Logger;

async function search(searchTerm, options) {
  const result = await sp.search({
    Querytext: options.exactMatch ? `"${searchTerm}"` : searchTerm,
    RowLimit: 10,
    EnableInterleaving: true
  });
  Logger.debug({ result: result.PrimarySearchResults }, 'Search Results');
  return result.PrimarySearchResults;
}

function isOptionChanged(options) {
  const newOptionsHash =
    options.onpremUsername + options.onpremPassword + options.onpremDomain + options.subsite + options.url;
  if (newOptionsHash !== optionsHash) {
    optionsHash = newOptionsHash;
    return true;
  }
  return false;
}

function setupSharePointLibrary({ url, onpremUsername: username, onpremPassword: password, onpremDomain: domain }) {
  const pnpNodeSettings = {
    siteUrl: url,
    authOptions: {
      username,
      password,
      domain
    }
  };

  Logger.trace({ pnpNodeSettings }, 'Sharepoint Client Settings');

  sp.setup({
    sp: {
      fetchClientFactory: () => {
        return new PnpNode(pnpNodeSettings);
      }
    }
  });
}

function _getSummaryTags(results) {
  const tags = [];
  results.forEach((result) => {
    if(result.FileExtension === 'aspx'){
      tags.push(`Page: ${result.Title}`);
    } else {
      tags.push(`File: ${result.Title}.${result.FileExtension}`);
    }
  });
  return tags;
}

function formatSearchResults(searchResults, options) {
  return searchResults.map((result) => {
    const formattedResult = { ...result };

    if (formattedResult.HitHighlightedSummary) {
      formattedResult.HitHighlightedSummary = formattedResult.HitHighlightedSummary.replace(/c0/g, 'strong').replace(
        /<ddd\/>/g,
        '&#8230;'
      );
    }

    if (formattedResult.FileType) {
      if (fileTypes[formattedResult.FileType]) {
        formattedResult._icon = fileTypes[formattedResult.FileType];
      } else {
        formattedResult._icon = 'file';
      }
    }

    if (formattedResult.Size) {
      formattedResult._sizeHumanReadable = xbytes(formattedResult.Size);
    }

    if (formattedResult.ParentLink) {
      formattedResult._containingFolder = url.parse(formattedResult.ParentLink).pathname;
    }

    return formattedResult;
  });
}

function errorToPojo(err, detail) {
  return err instanceof Error
    ? {
        ...err,
        name: err.name,
        message: err.message,
        stack: err.stack,
        detail: detail ? detail : err.detail ? err.detail : 'Unexpected error encountered'
      }
    : err;
}

async function doLookup(entities, options, cb) {
  const lookupResults = [];

  if (isOptionChanged(options)) {
    Logger.trace({ options }, 'Options Changed');
    setupSharePointLibrary(options);
  }

  try {
    await async.each(entities, async (entity) => {
      const searchResults = await search(entity.value, options);
      lookupResults.push({
        entity,
        data: {
          summary: _getSummaryTags(searchResults),
          details: formatSearchResults(searchResults, options)
        }
      });
    });
  } catch (lookupError) {
    Logger.error(lookupError, 'doLookupError');
    return cb(errorToPojo(lookupError, 'Failed to lookup entity'));
  }

  cb(null, lookupResults);
}

function startup(logger) {
  Logger = logger;
}

function validateStringOption(errors, options, optionName, errMessage) {
  if (
    typeof options[optionName].value !== 'string' ||
    (typeof options[optionName].value === 'string' && options[optionName].value.length === 0)
  ) {
    errors.push({
      key: optionName,
      message: errMessage
    });
  }
}

function validateOptions(options, callback) {
  let errors = [];

  validateStringOption(errors, options, 'url', 'You must provide a Sharepoint Site URL option.');
  validateStringOption(errors, options, 'onpremUsername', 'You must provide a Sharepoint Username.');
  validateStringOption(errors, options, 'onpremPassword', 'You must provide a password for the given username.');
  validateStringOption(errors, options, 'onpremDomain', 'You must provide a domain for the given username.');

  callback(null, errors);
}

module.exports = {
  doLookup: doLookup,
  startup: startup,
  validateOptions: validateOptions
};
