// vybe-test: js/map_set_iteration_groups/set_minus_set_operation_es2025
// origin: languages/js/tests/js/test_map_set_iteration_groups.rs

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

const a=new Set([1,2,3]); const b=new Set([2,3,4]); __check(__line(a.difference(b).size), "1");
