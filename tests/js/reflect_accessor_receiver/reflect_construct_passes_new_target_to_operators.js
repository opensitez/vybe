// vybe-test: js/reflect_accessor_receiver/reflect_construct_passes_new_target_to_operators
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

class A{constructor(){this.kind=new.target.name;}} class B extends A{} const i=Reflect.construct(B,[],A); __check(__line(i.kind), "A");
