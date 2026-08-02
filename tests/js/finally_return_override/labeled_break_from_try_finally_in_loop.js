// vybe-test: js/finally_return_override/labeled_break_from_try_finally_in_loop
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

let o=[];outer:for(let i=0;i<3;i++){try{o.push(i);if(i===1)throw i;}finally{if(i===1)break outer;}}console.log(o.join(","));
