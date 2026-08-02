// vybe-test: js/async_await_error_recovery/async_parallel_for_await_two_generators
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

async function main(){const o=[];async function* a(){yield "a";}async function* b(){throw "b";}const run=async(g,l)=>{try{for await(const v of g())o.push(l+v);}catch(e){o.push(l+"e:"+e);}};await Promise.all([run(a(),""),run(b(),"")]);console.log(o.sort().join(","));}main();
