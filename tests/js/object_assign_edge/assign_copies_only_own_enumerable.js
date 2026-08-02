// vybe-test: js/object_assign_edge/assign_copies_only_own_enumerable
// origin: languages/js/tests/js/test_object_assign_edge.rs

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
src.own = "yes";
Object.defineProperty(src, "hidden", { value: "no", enumerable: false });
const result = Object.assign({}, src);
__check(__line(result.own), "yes");
__check(__line(result.inherited), "undefined");
__check(__line(result.hidden), "undefined");
