// vybe-test: js/generator_async_delegate_errors/generator_delegate_to_custom_iterable_with_return
// origin: languages/js/tests/js/test_generator_async_delegate_errors.rs

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

const iterable={[Symbol.iterator](){let n=0;return{next(){return n++?{value:undefined,done:true}:{value:7,done:false};},return(v){return{value:v,done:true};}};}}; function* g(){const r=yield* iterable; yield r;} __check(__line([...g()][0]), "undefined");
