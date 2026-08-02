// vybe-test: js/throw_in_loops/infinite_while_caught_throw_counter
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

let o=[],n=0;while(n<5){try{if(n===3)throw n;}catch(e){o.push("c"+e);}o.push(n);n++;}console.log(o.join(","));
