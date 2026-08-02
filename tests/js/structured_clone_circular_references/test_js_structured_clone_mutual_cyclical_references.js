// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_mutual_cyclical_references
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

const nodeA = { id: "A" };
const nodeB = { id: "B" };
nodeA.sibling = nodeB;
nodeB.sibling = nodeA;

const cloneA = structuredClone(nodeA);
__check(__line((cloneA.sibling !== nodeB) + "|" + (cloneA.sibling.sibling === cloneA)), "true|true");
