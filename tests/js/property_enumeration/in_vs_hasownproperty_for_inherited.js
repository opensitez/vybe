// vybe-test: js/property_enumeration/in_vs_hasownproperty_for_inherited
// origin: languages/js/tests/js/test_property_enumeration.rs

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
const obj = Object.create(proto);
obj.own = 1;
__check(__line("own" in obj), "true");
__check(__line("inherited" in obj), "true");   // user-defined inherited property
__check(__line(obj.hasOwnProperty("own")), "true");
__check(__line(obj.hasOwnProperty("inherited")), "false");
