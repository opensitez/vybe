// vybe-test: js/try_catch_nested/nested_triple_instanceof_filter_inner
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
try{try{try{throw new TypeError("x");}catch(e){if(e instanceof TypeError){o.push("ok");}else throw e;}}
catch{o.push("bad");}}
catch{o.push("worse");}
__check(__line(o.join(",")), "ok");
