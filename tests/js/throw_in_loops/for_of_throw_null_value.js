// vybe-test: js/throw_in_loops/for_of_throw_null_value
// origin: languages/js/tests/js/test_throw_in_loops.rs

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

let o=[];try{for(const x of [1,null,3]){if(x===null)throw new TypeError("null");o.push(x);}}catch(e){o.push(e.name);}console.log(o.join(","));
