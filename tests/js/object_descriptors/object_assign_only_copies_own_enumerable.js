// vybe-test: js/object_descriptors/object_assign_only_copies_own_enumerable
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const proto = { inherited: 1 };
const src = Object.create(proto);
src.own = 2;
const result = Object.assign({}, src);
__check(__line(result.own), "2");
__check(__line(result.inherited), "undefined");
