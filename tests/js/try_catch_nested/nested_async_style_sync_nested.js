// vybe-test: js/try_catch_nested/nested_async_style_sync_nested
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
try{try{throw "sync";}catch(e){o.push("1:"+e);}}
catch(e){o.push("2:"+e);}
__check(__line(o.join(",")), "1:sync");
