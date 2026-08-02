// vybe-test: js/throw_in_loops/for_await_throw_from_async_iterator
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

async function main(){const o=[];async function* g(){yield 1;throw new Error("async");}try{for await(const v of g())o.push(v);}catch(e){o.push(e.message);}console.log(o.join(","));}main();
