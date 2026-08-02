// vybe-test: js/async_await_error_recovery/async_for_loop_await_with_switch_break
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

async function main(){const o=[];for(let i=0;i<3;i++){switch(i){case 0:try{await Promise.reject("s");}catch(e){o.push(e);}break;case 1:o.push("m");break;}}console.log(o.join(","));}main();
