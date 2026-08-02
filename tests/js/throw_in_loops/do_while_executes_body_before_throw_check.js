// vybe-test: js/throw_in_loops/do_while_executes_body_before_throw_check
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

let o=[],n=0;do{try{o.push(n);if(n===0)throw "once";}catch(e){o.push(String(e));}n++;}while(n<2);console.log(o.join(","));
