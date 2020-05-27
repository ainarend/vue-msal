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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const axios_1 = __importDefault(require("axios"));
const lodash_1 = __importDefault(require("lodash"));
const UserAgentApplicationExtended_1 = require("./UserAgentApplicationExtended");
class MSAL {
    constructor(options) {
        this.options = options;
        this.tokenExpirationTimer = undefined;
        this.data = {
            isAuthenticated: false,
            accessToken: '',
            user: {},
            userDetails: {},
            custom: {}
        };
        this.callbackQueue = [];
        this.auth = {
            clientId: '',
            tenantId: 'common',
            tenantName: 'login.microsoftonline.com',
            redirectUri: window.location.href,
            postLogoutRedirectUri: window.location.href,
            navigateToLoginRequestUrl: true,
            requireAuthOnInitialize: false,
            autoRefreshToken: true,
            onAuthentication: (error, response) => { },
            onToken: (error, response) => { },
            beforeSignOut: () => { }
        };
        this.cache = {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: true
        };
        this.request = {
            scopes: ["user.read"]
        };
        this.graph = {
            callAfterInit: false,
            meEndpoint: "https://graph.microsoft.com/v1.0/me",
            onResponse: (response) => { }
        };
        if (!options.auth.clientId) {
            throw new Error('auth.clientId is required');
        }
        this.auth = Object.assign(this.auth, options.auth);
        this.cache = Object.assign(this.cache, options.cache);
        this.request = Object.assign(this.request, options.request);
        this.graph = Object.assign(this.graph, options.graph);
        this.lib = new UserAgentApplicationExtended_1.UserAgentApplicationExtended({
            auth: {
                clientId: this.auth.clientId,
                authority: `https://${this.auth.tenantName}/${this.auth.tenantId}`,
                redirectUri: this.auth.redirectUri,
                postLogoutRedirectUri: this.auth.postLogoutRedirectUri,
                navigateToLoginRequestUrl: this.auth.navigateToLoginRequestUrl
            },
            cache: this.cache,
            system: options.system
        });
        this.getSavedCallbacks();
        this.executeCallbacks();
        // Register Callbacks for redirect flow
        this.lib.handleRedirectCallback((error, response) => {
            this.saveCallback('auth.onAuthentication', error, response);
        });
        this.lib.handleRedirectCallback((response) => {
            this.saveCallback('auth.onToken', null, response);
        }, (error) => {
            this.saveCallback('auth.onToken', error, null);
        });
        if (this.auth.requireAuthOnInitialize) {
            this.signIn();
        }
        this.data.isAuthenticated = this.isAuthenticated();
        if (this.data.isAuthenticated) {
            this.data.user = this.lib.getAccount();
            this.acquireToken().then(() => {
                if (this.graph.callAfterInit) {
                    this.callMSGraph();
                }
            });
        }
        this.getStoredCustomData();
    }
    signIn() {
        if (!this.lib.isCallback(window.location.hash) && !this.lib.getAccount()) {
            // request can be used for login or token request, however in more complex situations this can have diverging options
            this.lib.loginRedirect(this.request);
        }
    }
    signOut() {
        return __awaiter(this, void 0, void 0, function* () {
            if (this.options.auth.beforeSignOut) {
                yield this.options.auth.beforeSignOut(this);
            }
            this.lib.logout();
        });
    }
    isAuthenticated() {
        return !this.lib.isCallback(window.location.hash) && !!this.lib.getAccount();
    }
    acquireToken(request = this.request) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                //Always start with acquireTokenSilent to obtain a token in the signed in user from cache
                const { accessToken, expiresOn, scopes } = yield this.lib.acquireTokenSilent(request);
                this.setAccessToken(accessToken, expiresOn, scopes);
                return accessToken;
            }
            catch (error) {
                // Upon acquireTokenSilent failure (due to consent or interaction or login required ONLY)
                // Call acquireTokenRedirect
                if (this.requiresInteraction(error.errorCode)) {
                    this.lib.acquireTokenRedirect(request); //acquireTokenPopup
                }
                return false;
            }
        });
    }
    setAccessToken(accessToken, expiresOn, scopes) {
        this.data.accessToken = accessToken;
        // expiresOn = new Date(); expiresOn.setSeconds(expiresOn.getSeconds() + 10);
        // console.log('token set', expiresOn, expiresOn.getTime() - (new Date()).getTime(), scopes);
        const expiration = expiresOn.getTime() - (new Date()).getTime();
        if (this.tokenExpirationTimer)
            clearTimeout(this.tokenExpirationTimer);
        this.tokenExpirationTimer = setTimeout(() => {
            if (this.auth.autoRefreshToken) {
                this.acquireToken({ scopes });
            }
            else {
                this.data.accessToken = '';
            }
        }, expiration);
    }
    requiresInteraction(errorCode) {
        if (!errorCode || !errorCode.length) {
            return false;
        }
        return errorCode === "consent_required" ||
            errorCode === "interaction_required" ||
            errorCode === "login_required";
    }
    callMSGraph() {
        return __awaiter(this, void 0, void 0, function* () {
            const { onResponse: callback, meEndpoint } = this.graph;
            if (meEndpoint) {
                const storedData = this.lib.store.getItem(`msal.msgraph-${this.data.accessToken}`);
                if (storedData) {
                    this.data.userDetails = JSON.parse(storedData);
                }
                else {
                    try {
                        const response = yield axios_1.default.get(meEndpoint, {
                            headers: {
                                Authorization: 'Bearer ' + this.data.accessToken
                            }
                        });
                        this.data.userDetails = response.data;
                        this.lib.store.setItem(`msal.msgraph-${this.data.accessToken}`, JSON.stringify(this.data.userDetails));
                    }
                    catch (error) {
                        console.log(error);
                        return;
                    }
                }
                if (callback)
                    this.saveCallback('graph.onResponse', this.data.userDetails);
            }
        });
    }
    // CUSTOM DATA
    saveCustomData(key, data) {
        if (!this.data.custom.hasOwnProperty(key)) {
            this.data.custom[key] = null;
        }
        this.data.custom[key] = data;
        this.storeCustomData();
    }
    storeCustomData() {
        if (!lodash_1.default.isEmpty(this.data.custom)) {
            this.lib.store.setItem('msal.custom', JSON.stringify(this.data.custom));
        }
        else {
            this.lib.store.removeItem('msal.custom');
        }
    }
    getStoredCustomData() {
        let customData = {};
        const customDataStr = this.lib.store.getItem('msal.custom');
        if (customDataStr) {
            customData = JSON.parse(customDataStr);
        }
        this.data.custom = customData;
    }
    // CALLBACKS
    saveCallback(callbackPath, ...args) {
        if (lodash_1.default.get(this.options, callbackPath)) {
            const callbackQueueObject = {
                id: lodash_1.default.uniqueId(`cb-${callbackPath}`),
                callback: callbackPath,
                arguments: args
            };
            this.callbackQueue.push(callbackQueueObject);
            this.storeCallbackQueue();
            this.executeCallbacks([callbackQueueObject]);
        }
    }
    getSavedCallbacks() {
        const callbackQueueStr = this.lib.store.getItem('msal.callbackqueue');
        if (callbackQueueStr) {
            this.callbackQueue = [...this.callbackQueue, ...JSON.parse(callbackQueueStr)];
        }
    }
    executeCallbacks(callbacksToExec = this.callbackQueue) {
        return __awaiter(this, void 0, void 0, function* () {
            if (callbacksToExec.length) {
                for (let i in callbacksToExec) {
                    const cb = callbacksToExec[i];
                    const callback = lodash_1.default.get(this.options, cb.callback);
                    try {
                        yield callback(this, ...cb.arguments);
                        lodash_1.default.remove(this.callbackQueue, function (currentCb) {
                            return cb.id === currentCb.id;
                        });
                        this.storeCallbackQueue();
                    }
                    catch (e) {
                        console.warn(`Callback '${cb.id}' failed with error: `, e.message);
                    }
                }
            }
        });
    }
    storeCallbackQueue() {
        if (this.callbackQueue.length) {
            this.lib.store.setItem('msal.callbackqueue', JSON.stringify(this.callbackQueue));
        }
        else {
            this.lib.store.removeItem('msal.callbackqueue');
        }
    }
}
exports.MSAL = MSAL;
//# sourceMappingURL=main.js.map