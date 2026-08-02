// vybe-test: js/throw_in_loops/for_loop_continue_executes_finally
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

let o=[];for(let i=0;i<3;i++){try{if(i===1)continue;o.push(i);}finally{o.push("f"+i);}o.push("post");}console.log(o.join(","));
