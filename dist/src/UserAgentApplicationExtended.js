"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var msal_1 = require("msal");
var UserAgentApplicationExtended = /** @class */ (function (_super) {
    __extends(UserAgentApplicationExtended, _super);
    function UserAgentApplicationExtended(configuration) {
        var _this = _super.call(this, configuration) || this;
        _this.store = {};
        _this.store = _this.cacheStorage;
        return _this;
    }
    return UserAgentApplicationExtended;
}(msal_1.UserAgentApplication));
exports.UserAgentApplicationExtended = UserAgentApplicationExtended;
//# sourceMappingURL=UserAgentApplicationExtended.js.map