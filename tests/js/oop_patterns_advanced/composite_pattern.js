// vybe-test: js/oop_patterns_advanced/composite_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class File {
    constructor(name, size) { this.name=name; this.size=size; }
    totalSize() { return this.size; }
}
class Folder {
    constructor(name) { this.name=name; this.children=[]; }
    add(child) { this.children.push(child); return this; }
    totalSize() { return this.children.reduce((s,c)=>s+c.totalSize(),0); }
}
const root = new Folder("root")
    .add(new File("a.txt", 100))
    .add(new Folder("sub").add(new File("b.txt", 200)).add(new File("c.txt", 300)));
__check(__line(root.totalSize()), "600");
