// vybe-test: js/coercion_modern/object_to_map
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let obj = { x: 10, y: 20 };
let m = new Map(Object.entries(obj));
__check(__line(m.get("x")), "10");
__check(__line(m.get("y")), "20");
__check(__line(m.size), "2");
