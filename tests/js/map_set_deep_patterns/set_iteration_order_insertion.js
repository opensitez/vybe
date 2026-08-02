// vybe-test: js/map_set_deep_patterns/set_iteration_order_insertion
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

const s = new Set();
s.add(3); s.add(1); s.add(2); s.add(1); s.add(3);
__check(__line([...s].join(",")), "3,1,2");
__check(__line(s.size), "3");
