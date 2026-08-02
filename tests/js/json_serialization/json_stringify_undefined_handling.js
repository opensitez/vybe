// vybe-test: js/json_serialization/json_stringify_undefined_handling
// origin: languages/js/tests/js/test_json_serialization.rs

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

// JSON.stringify removes undefined values in objects
__check(__line(JSON.stringify({ a: undefined })), "{}");
// but keeps undefined in arrays as null
__check(__line(JSON.stringify([undefined, 1, undefined])), "[null,1,null]");
// standalone undefined returns undefined (not a string)
__check(__line(JSON.stringify(undefined)), "undefined");
