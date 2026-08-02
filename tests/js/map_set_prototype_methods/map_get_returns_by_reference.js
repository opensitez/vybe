// vybe-test: js/map_set_prototype_methods/map_get_returns_by_reference
// origin: languages/js/tests/js/test_map_set_prototype_methods.rs

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

const o={}; const m=new Map([["k",o]]); o.x=1; __check(__line(m.get("k").x), "1");
