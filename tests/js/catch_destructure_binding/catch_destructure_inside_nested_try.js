// vybe-test: js/catch_destructure_binding/catch_destructure_inside_nested_try
// origin: languages/js/tests/js/test_catch_destructure_binding.rs

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

let o=[];try{throw{layer:1};}catch({layer}){try{throw{layer:layer+1};}catch({layer:l}){o.push(l);}}__check(__line(o.join(",")), "2");
