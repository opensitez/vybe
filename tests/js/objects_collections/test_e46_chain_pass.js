// vybe-test: js/objects_collections/test_e46_chain_pass
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

function c(obj) { obj.z = 3; }
        function b(obj) { obj.y = 2; c(obj); }
        function a() { let o = { x: 1 }; b(o); return o; }
        let result = a();
        __check(__line(result.x, result.y, result.z), "1 2 3");
