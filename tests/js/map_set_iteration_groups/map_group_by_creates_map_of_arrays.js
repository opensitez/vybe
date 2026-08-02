// vybe-test: js/map_set_iteration_groups/map_group_by_creates_map_of_arrays
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

const g=Map.groupBy([1,2,3,4],n=>n%2===0?"even":"odd"); __check(__line(g.get("even").length), "2");
