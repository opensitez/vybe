// vybe-test: js/object_integrity_level_errors/freeze_array_blocks_push
// origin: languages/js/tests/js/test_object_integrity_level_errors.rs

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

const a=Object.freeze([1,2]); try{a.push(3); console.log("ok");}catch(e){console.log(e instanceof TypeError);}
