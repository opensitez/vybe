// vybe-test: js/try_catch_nested/try_inside_catch_preserves_outer_binding
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
try{throw {v:1};}
catch(e){try{o.push(e.v);}catch{o.push(0);}}
__check(__line(o.join(",")), "1");
