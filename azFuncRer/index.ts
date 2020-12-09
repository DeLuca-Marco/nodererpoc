// Credits for idea and implementation example:
// https://spblog.net/post/2017/09/09/Using-SharePoint-Remote-Event-Receivers-with-Azure-Functions-and-TypeScript

import { Parser } from "xml2js";
import { sp, SPRest } from "@pnp/sp-commonjs";
import { Web, IWeb } from "@pnp/sp-commonjs/webs";
import "@pnp/sp-commonjs/site-users/web";
import { NodeFetchClient, ProviderHostedRequestContext } from "@pnp/nodejs-commonjs";

declare var global: any;

export function run(context: any, req: any): void {
    context.log("Running RER from Azure Function");
    configurePnP();
    
    execute(context, req).catch((err: any) => {
        console.log(err);
        context.done();
    });
}

async function execute(context: any, req: any)
{
    let data = await xml2Json(req.body);
    context.log("running execute...");
    //Determine Event Method
    if(data["s:Envelope"]["s:Body"].ProcessOneWayEvent){
        await processOneWayEvent(data["s:Envelope"]["s:Body"].ProcessOneWayEvent.properties, context);
    } else if(data["s:Envelope"]["s:Body"].ProcessEvent){
        await processEvent(data["s:Envelope"]["s:Body"].ProcessEvent.properties, context)
    } else {
        throw new Error("Unable to resolve event type");
    }

}

async function processOneWayEvent(eventProperties: any, context: any): Promise<any> {
    context.log("running asyncrounous event");
    let itemProperties = eventProperties.ItemEventProperties;
    let spAddin = await getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
    let app = await spAddin.web.currentUser.get();
    context.log(app.Title);

    context.res = {
        status: 200,
        body: ""
    } as any;

    context.done();
}

async function processEvent(eventProperties: any, context: any): Promise<any> {
    context.log("running synchronous event");   
    let itemProperties = eventProperties.ItemEventProperties;
        let spAddin = await getAddinSP(eventProperties.ItemEventProperties.WebUrl, eventProperties.ContextToken);
        
        context.res = {
            status: 200,
            body: "",
        } as any;

        context.done();
}

function configurePnP(): void {
    // global.Headers = nodeFetch.Headers;
    // global.Request = nodeFetch.Request;
    // global.Response = nodeFetch.Response;

    sp.setup({
        sp: {
            fetchClientFactory: () => {
                return new NodeFetchClient();
            }
        }
    })
}

async function xml2Json(input: string): Promise<any> {
    return new Promise((resolve, reject) => {
        let parser = new Parser({
            explicitArray: false
        });

        parser.parseString(input, (jsError: any, jsResult: any) => {
            if (jsError) {
                reject(jsError);
            } else {
                resolve(jsResult);
            }
        });
    });
}

async function getUserSP(webUrl: string, contextToken: any) {
    let ctx = await initializeCtx(webUrl, contextToken);
    return new SPRest().configure(await ctx.getUserConfig(), webUrl);
}

async function getAddinSP(webUrl: string, contextToken: any) {
    let ctx = await initializeCtx(webUrl, contextToken);
    return new SPRest().configure(await ctx.getAddInOnlyConfig(), webUrl);
}

async function initializeCtx(webUrl: string, contextToken: any) {
    let spAppToken = contextToken;
    console.log(spAppToken);
    console.log(webUrl);
    return await ProviderHostedRequestContext.create(webUrl, getAppSettings("ClientId"), getAppSettings("ClientSecret"), spAppToken);
}

function getAppSettings(name: string): string {
    return process.env[name] as string;
}