// vybe-test: js/promise_rejection_propagation/parallel_rejections_independent_catches
// origin: languages/js/tests/js/test_promise_rejection_propagation.rs

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

const o=[];Promise.reject("a").catch(e=>o.push(e));Promise.reject("b").catch(e=>o.push(e));Promise.resolve().then(()=>console.log(o.sort().join(",")));
