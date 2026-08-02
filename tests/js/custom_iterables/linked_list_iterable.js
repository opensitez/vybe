// vybe-test: js/custom_iterables/linked_list_iterable
// origin: languages/js/tests/js/test_custom_iterables.rs

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

class Node {
    constructor(val, next = null) { this.val = val; this.next = next; }
}
class LinkedList {
    constructor() { this.head = null; }
    push(val) {
        const node = new Node(val);
        if (!this.head) { this.head = node; return; }
        let cur = this.head;
        while (cur.next) cur = cur.next;
        cur.next = node;
    }
    [Symbol.iterator]() {
        let cur = this.head;
        return {
            next() {
                if (cur) { const val = cur.val; cur = cur.next; return { value: val, done: false }; }
                return { done: true };
            }
        };
    }
}
const list = new LinkedList();
[1, 2, 3, 4, 5].forEach(x => list.push(x));
console.log([...list].join(","));
