// vybe-test: js/finally_return_override/loop_finally_break_exits_entire_loop
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

let o=[];for(let n=0;n<5;n++){try{o.push(n);if(n===2)throw n;}finally{if(n===2)break;}}console.log(o.join(","));
