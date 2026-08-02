// vybe-test: js/weakmap_weakset_patterns/weakset_tracks_objects
// origin: languages/js/tests/js/test_weakmap_weakset_patterns.rs

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

const seen = new WeakSet();
const a = {}, b = {}, c = {};
seen.add(a);
seen.add(b);
__check(__line(seen.has(a)), "true");
__check(__line(seen.has(c)), "false");
seen.delete(a);
__check(__line(seen.has(a)), "false");
