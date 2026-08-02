// vybe-test: js/reflect_accessor_receiver/reflect_apply_with_array_like_args
// origin: languages/js/tests/js/test_reflect_accessor_receiver.rs

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

function sum(a,b){return a+b;} __check(__line(Reflect.apply(sum,null,{0:3,1:4,length:2})), "7");
