// vybe-test: js/private_brand_errors/private_field_not_enumerable_in_reflect_ownkeys
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

class C{#h=1; a=2;} const k=Reflect.ownKeys(new C());__check(__line(k.includes("a")), "true");__check(__line(k.some(x=>typeof x==="symbol")), "false");
