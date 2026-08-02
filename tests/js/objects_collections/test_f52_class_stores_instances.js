// vybe-test: js/objects_collections/test_f52_class_stores_instances
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

class Item {
            constructor(name) { this.name = name; }
        }
        class Container {
            constructor() { this.items = []; }
            add(item) { this.items.push(item); }
            getName(i) { return this.items[i].name; }
        }
        let c = new Container();
        c.add(new Item("X"));
        c.add(new Item("Y"));
        __check(__line(c.getName(0), c.getName(1)), "X Y");
