// vybe-test: js/map_set_iteration_groups/weakmap_set_and_get_object_key
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

const wm=new WeakMap(); const k={}; wm.set(k,1); __check(__line(wm.get(k)), "1");
