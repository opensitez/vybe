// vybe-test: js/destructuring_advanced/object_destructure_prototype_not_included
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

const obj = Object.create({ inherited: true });
obj.own = "yes";
const { own, inherited } = obj;
__check(__line(own), "yes");
__check(__line(inherited), "true");
