/// Symbol.iterator custom implementations — pagination, linked list, tree, cyclic

use super::helpers::run_js;

#[test]
fn linked_list_iterable() {
    assert_eq!(run_js(r#"
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
"#), vec!["1,2,3,4,5"]);
}

#[test]
fn pagination_iterable() {
    assert_eq!(run_js(r#"
class Paginator {
    constructor(data, pageSize) { this.data = data; this.pageSize = pageSize; }
    [Symbol.iterator]() {
        let offset = 0;
        const { data, pageSize } = this;
        return {
            next() {
                if (offset >= data.length) return { done: true };
                const page = data.slice(offset, offset + pageSize);
                offset += pageSize;
                return { value: page, done: false };
            }
        };
    }
}
const pages = [...new Paginator([1,2,3,4,5,6,7], 3)];
console.log(pages.length);
console.log(pages[0].join(","));
console.log(pages[1].join(","));
console.log(pages[2].join(","));
"#), vec!["3", "1,2,3", "4,5,6", "7"]);
}

#[test]
fn cyclic_iterator_take() {
    assert_eq!(run_js(r#"
function* cycle(arr) {
    while (true) yield* arr;
}
function take(n, gen) {
    const result = [];
    for (const v of gen) { result.push(v); if (result.length >= n) break; }
    return result;
}
const colors = take(7, cycle(["red", "green", "blue"]));
console.log(colors.join(","));
"#), vec!["red,green,blue,red,green,blue,red"]);
}

#[test]
fn reverse_iterable() {
    assert_eq!(run_js(r#"
class ReverseIterable {
    constructor(arr) { this.arr = arr; }
    [Symbol.iterator]() {
        const arr = this.arr;
        let i = arr.length - 1;
        return {
            next() {
                return i >= 0 ? { value: arr[i--], done: false } : { done: true };
            }
        };
    }
}
const rev = new ReverseIterable([1, 2, 3, 4, 5]);
console.log([...rev].join(","));
"#), vec!["5,4,3,2,1"]);
}

#[test]
fn cartesian_product_generator() {
    assert_eq!(run_js(r#"
function* cartesian(a, b) {
    for (const x of a) for (const y of b) yield [x, y];
}
const pairs = [...cartesian([1, 2], ["a", "b"])];
console.log(pairs.length);
console.log(pairs.map(([x,y]) => x+y).join(","));
"#), vec!["4", "1a,1b,2a,2b"]);
}

#[test]
fn permutation_generator() {
    assert_eq!(run_js(r#"
function* permute(arr) {
    if (arr.length <= 1) { yield arr; return; }
    for (let i = 0; i < arr.length; i++) {
        const rest = [...arr.slice(0, i), ...arr.slice(i + 1)];
        for (const perm of permute(rest)) yield [arr[i], ...perm];
    }
}
const perms = [...permute([1, 2, 3])];
console.log(perms.length);
console.log(perms[0].join(","));
"#), vec!["6", "1,2,3"]);
}
