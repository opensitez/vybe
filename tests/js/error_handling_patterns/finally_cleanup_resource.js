// vybe-test: js/error_handling_patterns/finally_cleanup_resource
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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

const resources = [];
function openResource(id) {
    resources.push("open:" + id);
    return { id, close() { resources.push("close:" + id); } };
}
function process(id) {
    const res = openResource(id);
    try {
        if (id === 2) throw new Error("bad resource");
        return "ok";
    } finally {
        res.close();
    }
}
try { process(1); } catch {}
try { process(2); } catch {}
__check(__line(resources.join(",")), "open:1,close:1,open:2,close:2");
