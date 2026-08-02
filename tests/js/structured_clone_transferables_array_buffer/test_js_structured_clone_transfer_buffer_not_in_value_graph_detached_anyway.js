// vybe-test: js/structured_clone_transferables_array_buffer/test_js_structured_clone_transfer_buffer_not_in_value_graph_detached_anyway
// origin: languages/js/tests/js/test_js_structured_clone_transferables_array_buffer.rs

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

const bufToDetach = new ArrayBuffer(16);
const valueToClone = { msg: "Hello" };
const clone = structuredClone(valueToClone, { transfer: [bufToDetach] });

__check(__line(clone.msg + "|detached=" + (bufToDetach.byteLength === 0)), "Hello|detached=true");
