# PoC for a Node based RER

# Make your Azure Function available from the Internet
Use localtunnel for testing/developing, this command routes requests from the internet to your service:
lt -p 7071 -s vecrerpoc108

# Commonjs vs. ES6
You need to ensure that, by now (node 12.x - node 14.x) you only use commonjs libraries in your Azure Functions. Because node runs "server-side" JavaScript and therefore relies on commonjs it cannot handle ES6, which was developed for client-side JavaScript in the Browser.
So, errors like the following can be resolved by not using ES modules:
'Error [ERR_REQUIRE_ESM]: Must use import to load ES Module: C:\Users\deluc\proj\pocnoderer\node_modules\@pnp\sp\index.js
require() of ES modules is not supported...'

E.g. for PnPJS to work with node you have to use "@pnp/sp-commonjs" instead of "@pnp/sp" and "@pnp/nodejs-commonjs" instead of "@pnp/nodejs"

# Node Version for Azure Functions
Azure Functions currently works best with Node 12.

# Todos
Make synchronous events working. Therefore the SPRemoteEventResult object needs to be returned containing the ChangedItemProperties Dictionary. Therefore look what "AfterProperties" Object/Collection looks like.