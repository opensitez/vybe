// vybe-test: js/map_set_advanced_patterns/set_has_and_iteration_order
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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
s.add("c").add("a").add("b");
__check(__line([...s].join(",")), "c,a,b");
__check(__line(s.has("a")), "true");
s.delete("a");
__check(__line([...s].join(",")), "c,b");
