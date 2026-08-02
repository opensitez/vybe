// vybe-test: js/try_catch_nested/try_inside_catch_with_inner_finally_return
// origin: languages/js/tests/js/test_try_catch_nested.rs

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

let o=[];
function f(){ try{ throw 1; }catch(e){ try{ return "a"; }finally{ o.push("f"); } } }
f();
__check(__line(o.join(",")), "f");
