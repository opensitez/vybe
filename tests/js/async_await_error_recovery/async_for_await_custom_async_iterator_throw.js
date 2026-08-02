// vybe-test: js/async_await_error_recovery/async_for_await_custom_async_iterator_throw
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

async function main(){const o=[];const it={async next(){if(o.length)return{done:true};o.push("n");if(o.length===2)throw "n2";return{value:o.length,done:false};},[Symbol.asyncIterator](){return this;}};const r=[];try{for await(const v of it)r.push(v);}catch(e){r.push("e");}console.log(r.join(","));}main();
