/** 
 * Copyright 2023 Cloud Technology Solutions Ltd. All Rights Reserved.
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at

 *   http://www.apache.org/licenses/LICENSE-2.0

 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
**/

/** @interface **/

/**
 * Connect to an AppSheet App 
 * To enable the AppSheet API in your app:
 * 
 * 1.  Open the app in the app editor.
 * 2.  Select **Settings > Integrations**.
 * 3.  Under **IN: from cloud services to your app**, enable the **Enable** toggle.  
 *     This enables the API for the application as a whole.
 * 4.  Ensure that at least one unexpired **Application Access Key** is present.  
 *     Otherwise, click **Create Application Access Key**.
 * 5.  When you are done, save the app.
 * 6.  Use your app ID and Access Key to connect Apps Script to your app
 *
 * @param {String} appId AppSheet App ID.
 * @param {String} applicationAccessKey AppSheet App Access Key.
 * @return {AppSheetApp}
 */
function  connect(appId, applicationAccessKey) {
    return new AppSheetApp(appId, applicationAccessKey);
}

/**
 * Add records to a table
 * 
 * @param {String} tableName - specifies the name of the table
 * @param {Object[]} rows - One or more Rows elements. Each individual Row value must normally include the key field values of the record to be added. However, if the key field contains an Initial value, you can omit the key field value. For example, you should omit the key field value when the key field has an Initial value of UNIQUEID() or RANDBETWEEN(). The system will initialize the key field to the Initial value.
 * @param {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]    
 * @param {Boolean} isAsync - **Optional** return a AppSheet API request object instead of making an immediate API call
 *  @returns {Object} AppSheet Response
 */
function Add(tableName, rows, properties = {}, isAsync = false) {
    return isAsync
      ? AppSheetApp._appSheetParam(tableName, "Add", rows, properties)
      : AppSheetApp._appSheetAPI(tableName, "Add", rows, properties);
}

/**
 * Delete records from a table
 * 
 * @param {String} tableName - specifies the name of the table
 * @param {Object[]} rows - One or more Rows elements to be deleted. Each Row value may contain field values of the key field values of the record to be deleted.
 * @param {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]  
 * @param {Boolean} isAsync - **Optional** return a AppSheet API request object instead of making an immediate API call
 * @returns {Object} AppSheet Response
 */
function Delete(tableName, rows, properties = {}, isAsync = false) {
    return isAsync
      ? AppSheetApp._appSheetParam(tableName, "Delete", rows, properties)
      : AppSheetApp._appSheetAPI(tableName, "Delete", rows, properties);
}

/**
 * Update records in a table
 * 
 * @param {String} tableName - specifies the name of the table
 * @param {Object[]} rows - One or more Row values to be updated. Each individual Row value must include the key field values of the record to be updated.
 * @param {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]   
 * @param {Boolean} isAsync - **Optional** if true, return a AppSheet API request object instead of making an immediate API call
 * @returns {Object} AppSheet Response
 */
function Edit(tableName, rows, properties = {}, isAsync = false) {
    return isAsync
      ? AppSheetApp._appSheetParam(tableName, "Edit", rows, properties)
      : AppSheetApp._appSheetAPI(tableName, "Edit", rows, properties);
}

/**
 * Read records from a table
 * 
 * @param {String} tableName - specifies the name of the table
 * @param {Object[]} rows - **Optional**. You can omit the Selector property and specify input Rows containing the key values of the records to be read.
 * @param {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]. Additionally the optional `Selector` property can used to specify an expression to select and format the rows returned [[Ref](https://support.google.com/appsheet/answer/10105770#:~:text=Read-,selected%20rows,-In%20the%20Selector)].   
 * @param {Boolean} isAsync - **Optional** if true, return a AppSheet API request object instead of making an immediate API call
 * @returns {Object} AppSheet Response
 */
function Find(tableName, rows, properties = {}, isAsync = false) {
    return isAsync
      ? AppSheetApp._appSheetParam(tableName, "Find", rows, properties)
      : AppSheetApp._appSheetAPI(tableName, "Find", rows, properties);
}

/**
 * Invoke an action
 * 
 * @param {String} tableName - specifies the name of the table
 * @param {String} action - The action name.
 * @param {Object[]} rows - One or more Rows elements specifying the key field values of the rows to which the action is to be applied.
 * @param {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]   
 * @param {Boolean} isAsync - **Optional** if true, return a AppSheet API request object instead of making an immediate API call
 * @returns {Object} AppSheet Response
 */
function Action(tableName, action, rows, properties = {}, isAsync = false) {
    return isAsync
      ? AppSheetApp._appSheetParam(tableName, action , rows, properties)
      : AppSheetApp._appSheetAPI(tableName, action , rows, properties);
}

/**
 * @typedef {object} requestObject
 * @property {String} action - The action name.
 * @property {Object[]} rows - One or more Rows elements specifying the key field values of the rows to which the action is to be applied.
 * @property {Object} properties - **Optional**. Optional properties such as Locale, Location, Timezone, and UserId. [[Ref](https://support.google.com/appsheet/answer/10105398?hl=en#:~:text=for%20the%20table.-,Properties,-The%20properties%20of)]   
 * @property {Boolean} isAsync - **Optional** if true, return a AppSheet API request object instead of making an immediate API call
 */
/**
 * Makes multiple requests to the AppSheet API.
 *
 * @param {...requestObject} params - A variable number of objects, where each object represents a single AppSheet API request.
 * @returns {Object[]} An array of AppSheet API responses, with each response corresponding to a request in the order they were provided.
 */
function fetchAll(...params){
  try {
    const responses = UrlFetchApp.fetchAll(params);
    return responses.map(response => JSON.parse(response.getContentText()));
  } catch (e) {
  return e;
  }
}

/**
 * AppSheetApp is used to interface the AppSheet API.
 */
class _AppSheet {
    /** 
     * @constructor 
     * @param {String} appId - AppSheet Application ID.
     * @param {String} applicationAccessKey - AppSheet Access Key.
     * @return {_AppSheet}
    */
    constructor(appId, applicationAccessKey) {
        this.appId = appId;
        this.applicationAccessKey = applicationAccessKey;
    }

    Add(tableName, rows, properties = {}, isAsync = false ) {
      return isAsync
        ? this._appSheetParam(tableName, 'Add', rows, properties)
        : this._appSheetAPI(tableName, 'Add', rows, properties)
    }

    Delete(tableName, rows, properties = {},  isAsync = false ) {
      return isAsync
        ? this._appSheetParam(tableName, 'Delete', rows, properties)
        : this._appSheetAPI(tableName, 'Delete', rows, properties);
    }

    Edit(tableName, rows, properties = {}, isAsync = false) {
      return isAsync
        ? this._appSheetParam(tableName, 'Edit', rows, properties)
        : this._appSheetAPI(tableName, 'Edit', rows, properties);
    }

    Find(tableName, rows, properties = {}, isAsync = false) {
      return isAsync
        ? this._appSheetParam(tableName, 'Find', rows, properties)
        : this._appSheetAPI(tableName, 'Find', rows, properties);
    }

    Action(tableName, action, rows, properties = {}, isAsync = false) {
      return isAsync 
        ? this._appSheetParam(tableName,action, rows, properties)
        : this._appSheetAPI(tableName, action, rows, properties);
    }

    fetchAll(...params){
      try {
          const responses = UrlFetchApp.fetchAll(params);
          return responses.map(response => JSON.parse(response.getContentText()));
      } catch (e) {
          return e;
      }
    }

    _appSheetAPI(tableName, action, rows, properties) {
        const params = this._appSheetParam(tableName, action, rows, properties);
        const url = params.url;
        
        // Remove the url from params to avoid sending it in the request
        // This is necessary because UrlFetchApp does not support sending the url in the params object
        // and it expects the url to be a separate parameter.
        delete params.url; 

        try {
            const response = UrlFetchApp.fetch(url, params);
            return JSON.parse(response.getContentText());
        } catch (e) {
            return e;
      }
    }

    _appSheetParam(tableName, action, rows = [], properties = {}) {        
        return {
            url: `https://api.appsheet.com/api/v2/apps/${this.appId}/tables/${tableName}/Action`,
            method: 'post',
            contentType: 'application/json', 
            headers: { ApplicationAccessKey: this.applicationAccessKey }, 
            payload: JSON.stringify({ Action: action, Rows: rows, Properties: properties }), 
            muteHttpExceptions: true
        };
    }
}
var AppSheetApp = _AppSheet;