"use strict";
// Credits for idea and implementation example:
// https://spblog.net/post/2017/09/09/Using-SharePoint-Remote-Event-Receivers-with-Azure-Functions-and-TypeScript
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
        context.log("running oneway");
        let itemProperties = eventProperties.ItemEventProperties;
        let spAddin = yield getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
        let app = yield spAddin.web.currentUser.get();
        context.log(app.Title);
        context.res = {
            status: 200,
            body: ""
        };
        context.done();
    });
}
function processEvent(eventProperties, context) {
    return __awaiter(this, void 0, void 0, function* () {
        context.log("running event");
        let itemProperties = eventProperties.ItemEventProperties;
        let spAddin = yield getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
        let webTitle = yield (yield spAddin.web.get()).Title;
        context.log(webTitle);
        context.res = {
            status: 200,
            body: ""
        };
        context.done();
    });
}
function configurePnP() {
    // global.Headers = nodeFetch.Headers;
    // global.Request = nodeFetch.Request;
    // global.Response = nodeFetch.Response;
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
        console.log(spAppToken);
        console.log(webUrl);
        return yield nodejs_commonjs_1.ProviderHostedRequestContext.create(webUrl, getAppSettings("ClientId"), getAppSettings("ClientSecret"), spAppToken);
    });
}
function getAppSettings(name) {
    return process.env[name];
}
//# sourceMappingURL=index.js.map