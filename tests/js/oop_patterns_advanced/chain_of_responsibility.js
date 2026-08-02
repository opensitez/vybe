// vybe-test: js/oop_patterns_advanced/chain_of_responsibility
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class Handler {
    constructor(next = null) { this.next = next; }
    handle(req) { return this.next ? this.next.handle(req) : "unhandled"; }
}
class AuthHandler extends Handler {
    handle(req) { return req.auth ? super.handle(req) : "unauthorized"; }
}
class RateLimitHandler extends Handler {
    handle(req) { return req.rate > 100 ? "rate limited" : super.handle(req); }
}
class ResourceHandler extends Handler {
    handle(req) { return "ok:" + req.resource; }
}
const chain = new AuthHandler(new RateLimitHandler(new ResourceHandler()));
__check(__line(chain.handle({ auth: true, rate: 50, resource: "data" })), "ok:data");
__check(__line(chain.handle({ auth: false, rate: 50, resource: "data" })), "unauthorized");
__check(__line(chain.handle({ auth: true, rate: 200, resource: "data" })), "rate limited");
