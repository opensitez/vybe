// vybe-test: js/async_await_error_recovery/async_for_loop_continue_after_catch
// origin: languages/js/tests/js/test_async_await_error_recovery.rs

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

async function main(){const o=[];for(let i=0;i<3;i++){try{if(i===1)throw "skip";o.push(i);}catch{o.push("c");}}console.log(o.join(","));}main();
