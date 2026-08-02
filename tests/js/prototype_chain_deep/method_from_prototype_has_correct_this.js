// vybe-test: js/prototype_chain_deep/method_from_prototype_has_correct_this
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

const proto = {
    double() { return this.value * 2; }
};
const obj = Object.create(proto);
obj.value = 21;
__check(__line(obj.double()), "42");
