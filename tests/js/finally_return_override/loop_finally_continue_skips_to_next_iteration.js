// vybe-test: js/finally_return_override/loop_finally_continue_skips_to_next_iteration
// origin: languages/js/tests/js/test_finally_return_override.rs

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

let o=[];for(let n=0;n<4;n++){try{o.push("t"+n);if(n===1)throw n;}catch{}finally{if(n===1)continue;}o.push("a"+n);}console.log(o.join(","));
