// vybe-test: js/immutable_patterns/structural_sharing_via_object_create
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

const base = { common: "shared", own: "base" };
const variant = Object.create(base);
variant.own = "variant";
__check(__line(variant.common), "shared");  // from prototype
__check(__line(variant.own), "variant");     // own property
__check(__line(base.own), "base");        // original unchanged
