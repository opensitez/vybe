// vybe-test: js/prototype_chain_deep/own_property_shadows_inherited
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

const proto = { x: 1, toString() { return "proto"; } };
const obj = Object.create(proto);
obj.x = 99;
__check(__line(obj.x), "99");
__check(__line(proto.x), "1");
