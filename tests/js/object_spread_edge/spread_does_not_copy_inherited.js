// vybe-test: js/object_spread_edge/spread_does_not_copy_inherited
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const proto = { inherited: true };
const src = Object.create(proto);
src.own = true;
const result = { ...src };
__check(__line(result.own), "true");
__check(__line(result.inherited), "undefined");
