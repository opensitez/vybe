// vybe-test: js/try_catch_nested/nested_inner_finally_on_rethrow
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
try{
  try{
    try{ throw 1; }catch(e){ o.push("c"); throw e; }
  }finally{ o.push("f"); }
}catch(e){ o.push("o"); }
__check(__line(o.join(",")), "c,f,o");
