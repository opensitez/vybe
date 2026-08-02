// vybe-test: js/json_deep/stringify_circular_reference_throws
// origin: languages/js/tests/js/test_json_deep.rs

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

const obj = {};
obj.self = obj;
let threw = false;
try { JSON.stringify(obj); } catch (e) { threw = e instanceof TypeError; }
__check(__line(threw), "true");
