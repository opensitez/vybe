// vybe-test: js/map_set_iteration_groups/map_for_each_visits_all_entries
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

const o=[]; new Map([["a",1],["b",2]]).forEach((v,k)=>o.push(k+v)); console.log(o.sort().join(","));
