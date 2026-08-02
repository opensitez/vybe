// vybe-test: js/map_set_iteration_groups/map_spread_into_array_of_pairs
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

const p=[...new Map([["a",1]])][0]; __check(__line(p[0]), "a");__check(__line(p[1]), "1");
