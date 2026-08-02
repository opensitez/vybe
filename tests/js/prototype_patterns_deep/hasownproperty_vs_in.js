// vybe-test: js/prototype_patterns_deep/hasownproperty_vs_in
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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
const obj = Object.create(proto);
obj.own = 2;
__check(__line(obj.hasOwnProperty("own")), "true");
__check(__line(obj.hasOwnProperty("inherited")), "false");
__check(__line("own" in obj), "true");
__check(__line("inherited" in obj), "true");
