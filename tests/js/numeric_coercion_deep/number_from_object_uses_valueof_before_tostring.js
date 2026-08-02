// vybe-test: js/numeric_coercion_deep/number_from_object_uses_valueof_before_tostring
// origin: languages/js/tests/js/test_numeric_coercion_deep.rs

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

const obj = {
    order: [],
    valueOf() {
        this.order.push("valueOf");
        return 7;
    },
    toString() {
        this.order.push("toString");
        return "42";
    }
};

__check(__line(`${Number(obj)}|${obj.order.join(",")}`), "7|valueOf");
