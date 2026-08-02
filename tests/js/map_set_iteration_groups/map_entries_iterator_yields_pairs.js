// vybe-test: js/map_set_iteration_groups/map_entries_iterator_yields_pairs
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

const e=new Map([["k",9]]).entries().next().value; __check(__line(e[0]), "k");__check(__line(e[1]), "9");
