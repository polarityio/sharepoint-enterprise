const async = require('async');
const url = require('url');
const xbytes = require('xbytes');
const { sp } = require('@pnp/sp-commonjs');
const { PnpNode } = require('sp-pnp-node');

const MAX_TAGS = 5;
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
  json: 'file',
  log: 'file',
  aspx: 'microsoft'
};

let optionsHash = '';
let Logger;

async function search(searchTerm, options) {
  const result = await sp.search({
    Querytext: options.exactMatch ? `"${searchTerm}"` : searchTerm,
    RowLimit: 10,
    EnableInterleaving: true
  });
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
  // Use a set to dedupe tags
  const tags = new Set();
  results.forEach((result) => {
    if (result.FileExtension === 'aspx') {
      tags.add(`Page: ${result.Title}`);
    } else {
      tags.add(`File: ${result.Title}.${result.FileExtension}`);
    }
  });
  const tagList = [...tags];
  const slicedTagList = tagList.slice(0, MAX_TAGS);
  if (slicedTagList.length < tagList.length) {
    slicedTagList.push(`+${tagList.length - slicedTagList.length} results`);
  }
  return slicedTagList;
}

function formatSearchResults(searchResults, options) {
  return searchResults.reduce(
    (accum, result) => {
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

      if (formattedResult.FileExtension === 'aspx') {
        accum.pages.push(formattedResult);
      } else {
        accum.documents.push(formattedResult);
      }
      return accum;
    },
    {
      pages: [],
      documents: []
    }
  );
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
      if (searchResults.length === 0) {
        lookupResults.push({
          entity,
          data: null
        });
      } else {
        const formattedSearchResults = formatSearchResults(searchResults, options);
        Logger.debug({ result: formattedSearchResults }, 'Formatted Search Results');
        lookupResults.push({
          entity,
          data: {
            summary: _getSummaryTags(searchResults),
            details: formattedSearchResults
          }
        });
      }
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
