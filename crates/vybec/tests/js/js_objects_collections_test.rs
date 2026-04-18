use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

#[test]
fn test_a01_object_literal_with_properties() {
    let code = r#"
        let obj = { name: "Alice", age: 30 };
        console.log(obj.name, obj.age);
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

// A2. Object dynamic property add
#[test]
fn test_a02_object_dynamic_property_add() {
    let code = r#"
        let obj = { x: 1 };
        obj.y = 2;
        obj.z = 3;
        console.log(obj.x, obj.y, obj.z);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// A3. Nested objects: a.b.c access
#[test]
fn test_a03_nested_object_access() {
    let code = r#"
        let obj = { a: { b: { c: 99 } } };
        console.log(obj.a.b.c);
    "#;
    assert_eq!(run_js_one(code), "99");
}

// A4. Object passed to function — mutation visible to caller
#[test]
fn test_a04_object_pass_by_reference() {
    let code = r#"
        function modify(o) { o.val = 42; }
        let obj = { val: 0 };
        modify(obj);
        console.log(obj.val);
    "#;
    assert_eq!(run_js_one(code), "42");
}

// A5. Object returned from function
#[test]
fn test_a05_object_returned_from_function() {
    let code = r#"
        function make() { return { x: 10, y: 20 }; }
        let r = make();
        console.log(r.x, r.y);
    "#;
    assert_eq!(run_js_one(code), "10 20");
}

// A6. Object stored in array, retrieved, property accessed
#[test]
fn test_a06_object_in_array() {
    let code = r#"
        let arr = [{ name: "A" }, { name: "B" }, { name: "C" }];
        console.log(arr[1].name);
    "#;
    assert_eq!(run_js_one(code), "B");
}

// A7. Object spread: {...obj, extra: 1}
#[test]
fn test_a07_object_spread() {
    let code = r#"
        let base = { a: 1, b: 2 };
        let extended = { ...base, c: 3 };
        console.log(extended.a, extended.b, extended.c);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// A8. Object computed property: {[key]: value}
#[test]
fn test_a08_computed_property() {
    let code = r#"
        let key = "color";
        let obj = { [key]: "red" };
        console.log(obj.color);
    "#;
    assert_eq!(run_js_one(code), "red");
}

// A9. Object shorthand: {x, y} where x and y are variables
#[test]
fn test_a09_property_shorthand() {
    let code = r#"
        let x = 100;
        let y = 200;
        let obj = { x, y };
        console.log(obj.x, obj.y);
    "#;
    assert_eq!(run_js_one(code), "100 200");
}

// A10. Object identity: two vars same ref, modify through one
#[test]
fn test_a10_object_identity() {
    let code = r#"
        let a = { val: 1 };
        let b = a;
        b.val = 99;
        console.log(a.val);
    "#;
    assert_eq!(run_js_one(code), "99");
}

// A11. Object.keys, Object.values, Object.entries
#[test]
fn test_a11_object_keys_values_entries() {
    let code = r#"
        let obj = { x: 10 };
        let keys = Object.keys(obj);
        let vals = Object.values(obj);
        let entries = Object.entries(obj);
        console.log(keys.length, vals[0], entries[0][0], entries[0][1]);
    "#;
    assert_eq!(run_js_one(code), "1 10 x 10");
}

// A12. delete operator removes property
#[test]
fn test_a12_delete_property() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        delete obj.b;
        console.log("b" in obj, obj.a, obj.c);
    "#;
    assert_eq!(run_js_one(code), "false 1 3");
}

// ============================================================
// B. Map collection (8 tests)
// ============================================================

// B13. new Map() — set, get, has
#[test]
fn test_b13_map_set_get_has() {
    let code = r#"
        let m = new Map();
        m.set("name", "Alice");
        console.log(m.get("name"), m.has("name"), m.has("missing"));
    "#;
    assert_eq!(run_js_one(code), "Alice true false");
}

// B14. Map.size after multiple sets
#[test]
fn test_b14_map_size() {
    let code = r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.set("c", 3);
        console.log(m.size);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// B15. Map.delete
#[test]
fn test_b15_map_delete() {
    let code = r#"
        let m = new Map();
        m.set("x", 10);
        m.set("y", 20);
        m.delete("x");
        console.log(m.has("x"), m.size);
    "#;
    assert_eq!(run_js_one(code), "false 1");
}

// B16. Map.clear
#[test]
fn test_b16_map_clear() {
    let code = r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.clear();
        console.log(m.size);
    "#;
    assert_eq!(run_js_one(code), "0");
}

// B17. Map with string keys and number values
#[test]
fn test_b17_map_string_keys_number_values() {
    let code = r#"
        let m = new Map();
        m.set("score", 100);
        m.set("lives", 3);
        let total = m.get("score") + m.get("lives");
        console.log(total);
    "#;
    assert_eq!(run_js_one(code), "103");
}

// B18. Map — get missing key returns undefined
#[test]
fn test_b18_map_get_missing() {
    let code = r#"
        let m = new Map();
        console.log(m.get("nope"));
    "#;
    // callMethod returns null for missing keys (VM represents undefined/null as null)
    assert_eq!(run_js_one(code), "null");
}

// B19. Map — overwrite existing key
#[test]
fn test_b19_map_overwrite() {
    let code = r#"
        let m = new Map();
        m.set("key", "old");
        m.set("key", "new");
        console.log(m.get("key"), m.size);
    "#;
    assert_eq!(run_js_one(code), "new 1");
}

// B20. Map — multiple operations in sequence
#[test]
fn test_b20_map_operations_sequence() {
    let code = r#"
        let m = new Map();
        m.set("a", 1);
        m.set("b", 2);
        m.set("c", 3);
        m.delete("b");
        m.set("d", 4);
        console.log(m.size, m.has("a"), m.has("b"), m.get("d"));
    "#;
    assert_eq!(run_js_one(code), "3 true false 4");
}

// ============================================================
// C. Set collection (6 tests)
// ============================================================

// C21. new Set() — add, has
#[test]
fn test_c21_set_add_has() {
    let code = r#"
        let s = new Set();
        s.add("hello");
        s.add("world");
        console.log(s.has("hello"), s.has("missing"));
    "#;
    assert_eq!(run_js_one(code), "true false");
}

// C22. Set.size — duplicates not counted
#[test]
fn test_c22_set_size_no_duplicates() {
    let code = r#"
        let s = new Set();
        s.add("a");
        s.add("b");
        s.add("a");
        s.add("c");
        s.add("b");
        console.log(s.size);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// C23. Set.delete
#[test]
fn test_c23_set_delete() {
    let code = r#"
        let s = new Set();
        s.add(10);
        s.add(20);
        s.delete(10);
        console.log(s.has(10), s.size);
    "#;
    assert_eq!(run_js_one(code), "false 1");
}

// C24. Set.clear
#[test]
fn test_c24_set_clear() {
    let code = r#"
        let s = new Set();
        s.add(1);
        s.add(2);
        s.clear();
        console.log(s.size);
    "#;
    assert_eq!(run_js_one(code), "0");
}

// C25. Set — add same value twice, size stays 1
#[test]
fn test_c25_set_add_duplicate() {
    let code = r#"
        let s = new Set();
        s.add(42);
        s.add(42);
        console.log(s.size);
    "#;
    assert_eq!(run_js_one(code), "1");
}

// C26. Set — has returns false for missing
#[test]
fn test_c26_set_has_missing() {
    let code = r#"
        let s = new Set();
        s.add(1);
        console.log(s.has(999));
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// D. Array deep operations (12 tests)
// ============================================================

// D27. Array literal, push, length
#[test]
fn test_d27_array_push_length() {
    let code = r#"
        let arr = [1, 2];
        arr.push(3);
        arr.push(4);
        console.log(arr.length, arr.join(","));
    "#;
    assert_eq!(run_js_one(code), "4 1,2,3,4");
}

// D28. Array.map returning new array
#[test]
fn test_d28_array_map() {
    let code = r#"
        let doubled = [1, 2, 3].map(x => x * 2);
        console.log(doubled.join(","));
    "#;
    assert_eq!(run_js_one(code), "2,4,6");
}

// D29. Array.filter
#[test]
fn test_d29_array_filter() {
    let code = r#"
        let evens = [1, 2, 3, 4, 5, 6].filter(x => x % 2 === 0);
        console.log(evens.join(","));
    "#;
    assert_eq!(run_js_one(code), "2,4,6");
}

// D30. Array.reduce to sum
#[test]
fn test_d30_array_reduce_sum() {
    let code = r#"
        let sum = [1, 2, 3, 4, 5].reduce((acc, x) => acc + x, 0);
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "15");
}

// D31. Array.find
#[test]
fn test_d31_array_find() {
    let code = r#"
        let found = [10, 20, 30, 40].find(x => x > 25);
        console.log(found);
    "#;
    assert_eq!(run_js_one(code), "30");
}

// D32. Array.some / Array.every
#[test]
fn test_d32_array_some_every() {
    let code = r#"
        console.log([1, 2, 3].some(x => x > 2));
        console.log([2, 4, 6].every(x => x % 2 === 0));
        console.log([2, 3, 6].every(x => x % 2 === 0));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "true");
    assert_eq!(lines[1], "true");
    assert_eq!(lines[2], "false");
}

// D33. Array.findIndex
#[test]
fn test_d33_array_find_index() {
    let code = r#"
        console.log([10, 20, 30].findIndex(x => x === 20));
        console.log([10, 20, 30].findIndex(x => x === 99));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "1");
    assert_eq!(lines[1], "-1");
}

// D34. Array.sort with comparator
#[test]
fn test_d34_array_sort_comparator() {
    let code = r#"
        let arr = [3, 1, 4, 1, 5, 9];
        arr.sort((a, b) => a - b);
        console.log(arr.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,1,3,4,5,9");
}

// D35. Array.reverse (mutates)
#[test]
fn test_d35_array_reverse() {
    let code = r#"
        let arr = [1, 2, 3, 4];
        arr.reverse();
        console.log(arr.join(","));
    "#;
    assert_eq!(run_js_one(code), "4,3,2,1");
}

// D36. Array.concat (new array)
#[test]
fn test_d36_array_concat() {
    let code = r#"
        let a = [1, 2];
        let b = [3, 4];
        let c = a.concat(b);
        console.log(c.join(","), a.length);
    "#;
    assert_eq!(run_js_one(code), "1,2,3,4 2");
}

// D37. Array.slice (non-mutating)
#[test]
fn test_d37_array_slice() {
    let code = r#"
        let arr = [10, 20, 30, 40, 50];
        let sliced = arr.slice(1, 4);
        console.log(sliced.join(","), arr.length);
    "#;
    assert_eq!(run_js_one(code), "20,30,40 5");
}

// D38. Array.flat (one level)
#[test]
fn test_d38_array_flat() {
    let code = r#"
        let arr = [1, [2, 3], [4]];
        console.log(arr.flat().join(","));
    "#;
    assert_eq!(run_js_one(code), "1,2,3,4");
}

// D39. Array.fill
#[test]
fn test_d39_array_fill() {
    let code = r#"
        let a = [1, 2, 3, 4];
        a.fill(0, 1, 3);
        console.log(a.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,0,0,4");
}

// D40. Array.join with separator
#[test]
fn test_d40_array_join() {
    let code = r#"
        let arr = ["hello", "world", "foo"];
        console.log(arr.join(" - "));
    "#;
    assert_eq!(run_js_one(code), "hello - world - foo");
}

// ============================================================
// E. Passing objects (8 tests)
// ============================================================

// E41. Function modifies object param — caller sees change
#[test]
fn test_e41_function_modifies_object() {
    let code = r#"
        function inc(obj) { obj.count = obj.count + 1; }
        let o = { count: 0 };
        inc(o);
        inc(o);
        inc(o);
        console.log(o.count);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// E42. Function returns new object
#[test]
fn test_e42_function_returns_new_object() {
    let code = r#"
        function makePoint(x, y) { return { x: x, y: y }; }
        let p = makePoint(3, 4);
        console.log(p.x, p.y);
    "#;
    assert_eq!(run_js_one(code), "3 4");
}

// E43. Function takes array param, pushes item
#[test]
fn test_e43_function_pushes_to_array() {
    let code = r#"
        function addItem(arr, item) { arr.push(item); }
        let list = [1, 2];
        addItem(list, 3);
        addItem(list, 4);
        console.log(list.join(","));
    "#;
    assert_eq!(run_js_one(code), "1,2,3,4");
}

// E44. Object with method passed to function, function calls method
#[test]
fn test_e44_object_method_called_in_function() {
    let code = r#"
        let obj = { greet: function() { return "hello"; } };
        function callGreet(o) { return o.greet(); }
        console.log(callGreet(obj));
    "#;
    assert_eq!(run_js_one(code), "hello");
}

// E45. Recursive function building array
#[test]
fn test_e45_recursive_array_build() {
    let code = r#"
        function range(n) {
            if (n <= 0) return [];
            let arr = range(n - 1);
            arr.push(n);
            return arr;
        }
        console.log(range(5).join(","));
    "#;
    assert_eq!(run_js_one(code), "1,2,3,4,5");
}

// E46. Chain: A creates obj, passes to B, B passes to C
#[test]
fn test_e46_chain_pass() {
    let code = r#"
        function c(obj) { obj.z = 3; }
        function b(obj) { obj.y = 2; c(obj); }
        function a() { let o = { x: 1 }; b(o); return o; }
        let result = a();
        console.log(result.x, result.y, result.z);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// E47. Callback receives object as argument
#[test]
fn test_e47_callback_with_object() {
    let code = r#"
        function process(obj, cb) { cb(obj); }
        let data = { val: 10 };
        process(data, function(o) { o.val = o.val * 2; });
        console.log(data.val);
    "#;
    assert_eq!(run_js_one(code), "20");
}

// E48. Closure captures object, modifies it
#[test]
fn test_e48_closure_captures_object() {
    let code = r#"
        let obj = { count: 0 };
        let increment = () => { obj.count = obj.count + 1; };
        increment();
        increment();
        increment();
        console.log(obj.count);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// ============================================================
// F. Class with collections (8 tests)
// ============================================================

// F49. Class with array field — push in method
#[test]
fn test_f49_class_array_field_push() {
    let code = r#"
        class Bag {
            constructor() { this.items = []; }
            add(item) { this.items.push(item); }
            count() { return this.items.length; }
        }
        let b = new Bag();
        b.add("apple");
        b.add("banana");
        console.log(b.count());
    "#;
    assert_eq!(run_js_one(code), "2");
}

// F50. Class with Map field — set/get in methods
#[test]
fn test_f50_class_map_field() {
    let code = r#"
        class Registry {
            constructor() { this.data = new Map(); }
            register(key, val) { this.data.set(key, val); }
            lookup(key) { return this.data.get(key); }
        }
        let r = new Registry();
        r.register("host", "localhost");
        console.log(r.lookup("host"));
    "#;
    assert_eq!(run_js_one(code), "localhost");
}

// F51. Class method returns array
#[test]
fn test_f51_class_method_returns_array() {
    let code = r#"
        class NumGen {
            constructor(n) { this.n = n; }
            generate() {
                let arr = [];
                let i = 0;
                while (i < this.n) { arr.push(i); i = i + 1; }
                return arr;
            }
        }
        let g = new NumGen(4);
        console.log(g.generate().join(","));
    "#;
    assert_eq!(run_js_one(code), "0,1,2,3");
}

// F52. Class stores other class instances in array
#[test]
fn test_f52_class_stores_instances() {
    let code = r#"
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
        console.log(c.getName(0), c.getName(1));
    "#;
    assert_eq!(run_js_one(code), "X Y");
}

// F53. Class with constructor taking array param
#[test]
fn test_f53_class_constructor_array_param() {
    let code = r#"
        class Holder {
            constructor(data) { this.data = data; }
            first() { return this.data[0]; }
            size() { return this.data.length; }
        }
        let h = new Holder([10, 20, 30]);
        console.log(h.first(), h.size());
    "#;
    assert_eq!(run_js_one(code), "10 3");
}

// F54. Class iterating own array field
#[test]
fn test_f54_class_iterate_array_field() {
    let code = r#"
        class Summer {
            constructor() { this.values = []; }
            add(v) { this.values.push(v); }
            total() {
                let s = 0;
                let i = 0;
                while (i < this.values.length) {
                    s = s + this.values[i];
                    i = i + 1;
                }
                return s;
            }
        }
        let sm = new Summer();
        sm.add(10);
        sm.add(20);
        sm.add(30);
        console.log(sm.total());
    "#;
    assert_eq!(run_js_one(code), "60");
}

// F55. Two instances with independent array fields
#[test]
fn test_f55_independent_instances() {
    let code = r#"
        class Stack {
            constructor() { this.items = []; }
            push(v) { this.items.push(v); }
            size() { return this.items.length; }
        }
        let a = new Stack();
        let b = new Stack();
        a.push(1);
        a.push(2);
        b.push(99);
        console.log(a.size(), b.size());
    "#;
    assert_eq!(run_js_one(code), "2 1");
}

// F56. Class method modifying shared reference
#[test]
fn test_f56_class_shared_reference() {
    let code = r#"
        class Modifier {
            constructor(obj) { this.obj = obj; }
            setVal(v) { this.obj.val = v; }
        }
        let shared = { val: 0 };
        let m1 = new Modifier(shared);
        let m2 = new Modifier(shared);
        m1.setVal(10);
        console.log(shared.val);
        m2.setVal(20);
        console.log(shared.val);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "10");
    assert_eq!(lines[1], "20");
}

// ============================================================
// G. Iteration patterns (8 tests)
// ============================================================

// G57. for...of over array
#[test]
fn test_g57_for_of_array() {
    let code = r#"
        let sum = 0;
        for (let x of [10, 20, 30]) {
            sum = sum + x;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "60");
}

// G58. for...in over object keys
#[test]
fn test_g58_for_in_object() {
    let code = r#"
        let obj = { a: 1, b: 2, c: 3 };
        let keys = [];
        for (let k in obj) {
            keys.push(k);
        }
        console.log(keys.length);
    "#;
    assert_eq!(run_js_one(code), "3");
}

// G59. while loop processing array
#[test]
fn test_g59_while_loop_array() {
    let code = r#"
        let arr = [5, 10, 15, 20];
        let i = 0;
        let sum = 0;
        while (i < arr.length) {
            sum = sum + arr[i];
            i = i + 1;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "50");
}

// G60. Array.forEach
#[test]
fn test_g60_array_foreach() {
    let code = r#"
        let result = [];
        [1, 2, 3].forEach(x => { result.push(x * 10); });
        console.log(result.join(","));
    "#;
    assert_eq!(run_js_one(code), "10,20,30");
}

// G61. Nested for loops over 2D array
#[test]
fn test_g61_nested_loops_2d_array() {
    let code = r#"
        let grid = [[1, 2], [3, 4], [5, 6]];
        let sum = 0;
        let i = 0;
        while (i < grid.length) {
            let j = 0;
            while (j < grid[i].length) {
                sum = sum + grid[i][j];
                j = j + 1;
            }
            i = i + 1;
        }
        console.log(sum);
    "#;
    assert_eq!(run_js_one(code), "21");
}

// G62. Chained: arr.filter().map().join()
#[test]
fn test_g62_chained_array_methods() {
    let code = r#"
        let result = [1, 2, 3, 4, 5, 6]
            .filter(x => x % 2 === 0)
            .map(x => x * 10)
            .join(",");
        console.log(result);
    "#;
    assert_eq!(run_js_one(code), "20,40,60");
}

// G63. Building object from array via reduce
#[test]
fn test_g63_reduce_to_object() {
    let code = r#"
        let pairs = [["a", 1], ["b", 2], ["c", 3]];
        let obj = pairs.reduce((acc, pair) => {
            acc[pair[0]] = pair[1];
            return acc;
        }, {});
        console.log(obj.a, obj.b, obj.c);
    "#;
    assert_eq!(run_js_one(code), "1 2 3");
}

// G64. for...of with index using manual counter
#[test]
fn test_g64_for_of_with_counter() {
    let code = r#"
        let arr = ["x", "y", "z"];
        let result = [];
        let idx = 0;
        for (let item of arr) {
            result.push(idx + ":" + item);
            idx = idx + 1;
        }
        console.log(result.join(","));
    "#;
    assert_eq!(run_js_one(code), "0:x,1:y,2:z");
}
