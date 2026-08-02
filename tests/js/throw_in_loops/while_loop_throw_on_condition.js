// vybe-test: js/throw_in_loops/while_loop_throw_on_condition
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

let o=[],n=0;while(n<3){try{if(n===1)throw new RangeError("w");o.push(n);}catch(e){o.push(e.name);}n++;}console.log(o.join(","));
