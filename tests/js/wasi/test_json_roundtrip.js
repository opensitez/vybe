// vybe-test: js/wasi/test_json_roundtrip
// origin: languages/js/tests/js/js_wasi_test.rs

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

let obj = { x: 10, y: "hello" };
        let json = JSON.stringify(obj);
        let parsed = JSON.parse(json);
        __check(__line(parsed.x, parsed.y), "10 hello");
