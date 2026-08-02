// vybe-test: js/function_prototype_deep/call_extracted_method_without_receiver_loses_this
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

const obj = { n: 3, read() { return this.n; } }; const bare = obj.read; try { bare.call(null); console.log("ok"); } catch (e) { console.log(e instanceof TypeError); }
