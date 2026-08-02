// vybe-test: js/symbol_advanced/symbol_as_private_like_key
// origin: languages/js/tests/js/test_symbol_advanced.rs

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

const _private = Symbol("private");
function makeObj(secret) {
    const obj = {};
    obj[_private] = secret;
    obj.getSecret = function() { return this[_private]; };
    return obj;
}
const o = makeObj("shhh");
__check(__line(o.getSecret()), "shhh");
__check(__line(o[_private]), "shhh");
__check(__line(Object.keys(o).join(",")), "getSecret");
