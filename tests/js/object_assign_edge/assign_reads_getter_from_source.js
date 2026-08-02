// vybe-test: js/object_assign_edge/assign_reads_getter_from_source
// origin: languages/js/tests/js/test_object_assign_edge.rs

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

// assign copies own enumerable properties; test with nested reference
const nested = { deep: 42 };
const src = { ref: nested, str: "ok" };
const result = Object.assign({}, src);
__check(__line(result.ref === nested), "true");
__check(__line(result.str), "ok");
