"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.run = void 0;
// Credits for idea and implementation example:
// https://spblog.net/post/2017/09/09/Using-SharePoint-Remote-Event-Receivers-with-Azure-Functions-and-TypeScript
const fs = require("fs");
const xml2js_1 = require("xml2js");
const sp_commonjs_1 = require("@pnp/sp-commonjs");
require("@pnp/sp-commonjs/site-users/web");
const nodejs_commonjs_1 = require("@pnp/nodejs-commonjs");
function run(context, req) {
    context.log("Running RER from Azure Function");
    configurePnP();
    execute(context, req).catch((err) => {
        console.log(err);
        context.done();
    });
}
exports.run = run;
function execute(context, req) {
    return __awaiter(this, void 0, void 0, function* () {
        let data = yield xml2Json(req.body);
        context.log("running execute...");
        //Determine Event Method
        if (data["s:Envelope"]["s:Body"].ProcessOneWayEvent) {
            yield processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
        }
        else if (data["s:Envelope"]["s:Body"].ProcessEvent) {
            yield processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context);
        }
        else {
            throw new Error("Unable to resolve event type");
        }
    });
}
function processOneWayEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        context.log("running asyncrounous event");
        //constants
        const defaultGroupName = "Besitzer von Mod Spielwiese";
        const fieldNamePermissionIndicator = "permTag";
        let foundGroups = new Array();
        let itemProperties = eventProperties.ItemEventProperties;
        let spAddin = yield getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
        //Break files Role inheritance and set custom permissions
        let defaultGroup = yield GetGroupByName(spAddin, defaultGroupName);
        let item = yield spAddin.web.lists.getByTitle(eventProperties.ItemEventProperties.ListTitle).items.getById(eventProperties.ItemEventProperties.ListItemId);
        let fieldValues = yield item.fieldValuesAsText();
        let foundgroupNames = fieldValues[fieldNamePermissionIndicator].split(", ");
        for (let groupName of foundgroupNames) {
            context.log(groupName);
            foundGroups.push(yield GetGroupByName(spAddin, groupName + " Users"));
        }
        yield item.breakRoleInheritance(false);
        //add default group
        yield item.roleAssignments.add(defaultGroup.Id, 1073741829); //Full Control
        for (let foundGroup of foundGroups) {
            //add group defined by metadata
            yield item.roleAssignments.add(foundGroup.Id, 1073741826); //Read
        }
        context.res = {
            status: 200,
            body: ""
        };
        context.done();
    });
}
function processEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        let processResponse = initializeEventResponse();
        context.log("running synchronous event");
        //currently not necessary in this szenario, but shown examplary
        let spAddin = yield getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
        processResponse = AddChangedItemProperties(processResponse, "Title", "Changed by RER");
        processResponse = finalizeEventResponse(processResponse, "Continue", "", eventProperties.CorrelationId);
        context.res = {
            status: 200,
            headers: {
                "Content-Type": "text/xml; charset=utf-8"
            },
            body: processResponse,
            isRaw: true
        };
        context.log(context);
        context.done();
    });
}
function GetGroupByName(reqObj, name) {
    return __awaiter(this, void 0, void 0, function* () {
        return yield reqObj.web.siteGroups.getByName(name).get();
    });
}
function initializeEventResponse() {
    return fs.readFileSync("azFuncRer/response.data").toString();
}
function finalizeEventResponse(processResponse, status, errorMessage, correlationId) {
    if (!errorMessage) {
        processResponse = processResponse.replace(">###ERRORMESSAGE###", " i:nil=\"true\"/>").replace("</ErrorMessage>", "");
    }
    return processResponse.replace("###STATUS###", status).replace("###ERRORMESSAGE###", errorMessage).replace("###CORRELATIONID###", correlationId);
}
function AddChangedItemProperties(processResponse, fieldName, fieldValue) {
    let itemTempalte = fs.readFileSync("azFuncRer/changedItemProperties.data").toString();
    return processResponse.replace("</ChangedItemProperties>", itemTempalte.replace("###KEY###", fieldName).replace("###VALUE###", fieldValue) + "</ChangedItemProperties>");
}
function configurePnP() {
    sp_commonjs_1.sp.setup({
        sp: {
            fetchClientFactory: () => {
                return new nodejs_commonjs_1.NodeFetchClient();
            }
        }
    });
}
function xml2Json(input) {
    return __awaiter(this, void 0, void 0, function* () {
        return new Promise((resolve, reject) => {
            let parser = new xml2js_1.Parser({
                explicitArray: false
            });
            parser.parseString(input, (jsError, jsResult) => {
                if (jsError) {
                    reject(jsError);
                }
                else {
                    resolve(jsResult);
                }
            });
        });
    });
}
function getUserSP(webUrl, contextToken) {
    return __awaiter(this, void 0, void 0, function* () {
        let ctx = yield initializeCtx(webUrl, contextToken);
        return new sp_commonjs_1.SPRest().configure(yield ctx.getUserConfig(), webUrl);
    });
}
function getAddinSP(webUrl, contextToken) {
    return __awaiter(this, void 0, void 0, function* () {
        let ctx = yield initializeCtx(webUrl, contextToken);
        return new sp_commonjs_1.SPRest().configure(yield ctx.getAddInOnlyConfig(), webUrl);
    });
}
function initializeCtx(webUrl, contextToken) {
    return __awaiter(this, void 0, void 0, function* () {
        let spAppToken = contextToken;
        console.log(webUrl);
        return yield nodejs_commonjs_1.ProviderHostedRequestContext.create(webUrl, getAppSettings("ClientId"), getAppSettings("ClientSecret"), spAppToken);
    });
}
function getAppSettings(name) {
    return process.env[name];
}
//# sourceMappingURL=index.js.map