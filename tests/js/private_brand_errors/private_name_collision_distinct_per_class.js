// vybe-test: js/private_brand_errors/private_name_collision_distinct_per_class
// origin: languages/js/tests/js/test_private_brand_errors.rs

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

class A{#x(){return "a";}} class B{#x(){return "b";}} __check(__line(new A().#x()), "a");__check(__line(new B().#x()), "b");
