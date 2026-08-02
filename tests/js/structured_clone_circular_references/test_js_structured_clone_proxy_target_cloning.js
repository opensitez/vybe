// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_proxy_target_cloning
// origin: languages/js/tests/js/test_js_structured_clone_circular_references.rs

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

const target = { x: 50 };
const proxy = new Proxy(target, {});
const clone = structuredClone(proxy);
__check(__line(clone.x + "|" + (clone !== target)), "50|true");
