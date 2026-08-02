// vybe-test: js/weakref_weakmap_advanced/weakset_seen_objects_dedup_pattern
// origin: languages/js/tests/js/test_weakref_weakmap_advanced.rs

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
function process(obj) {
  if (seen.has(obj)) return "duplicate";
  seen.add(obj);
  return "new";
}
const a = {};
__check(__line(process(a)), "new");
__check(__line(process(a)), "duplicate");
__check(__line(process({})), "new");
