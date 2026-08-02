// vybe-test: js/ecma_operators/optional_chaining_call_eager_property_lookup_only
// origin: languages/js/tests/js/test_ecma_operators.rs

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

let calls = 0;
const target = {
    get printer() {
        calls += 1;
        return () => {
            calls += 10;
            return calls;
        };
    }
};

__check(__line(target?.printer?.()), "11");
__check(__line((null?.printer)?.()), "undefined");
__check(__line(calls), "11");
