// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_proxy_target_inspection
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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

class Token {
    #tokenVal = "ABC";
    static isToken(obj) { return #tokenVal in obj; }
}
const t = new Token();
const proxy = new Proxy(t, {});
__check(__line(Token.isToken(t) + "|" + Token.isToken(proxy)), "true|true"); // Brand check on Proxy succeeds if target is instance
