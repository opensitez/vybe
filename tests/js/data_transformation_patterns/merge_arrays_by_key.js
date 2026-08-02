// vybe-test: js/data_transformation_patterns/merge_arrays_by_key
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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

function mergeBy(key, ...arrays) {
    const map = new Map();
    for (const arr of arrays) {
        for (const item of arr) {
            const k = item[key];
            map.set(k, { ...(map.get(k) ?? {}), ...item });
        }
    }
    return [...map.values()];
}
const names = [{ id: 1, name: "Alice" }, { id: 2, name: "Bob" }];
const ages = [{ id: 1, age: 30 }, { id: 2, age: 25 }];
const merged = mergeBy("id", names, ages);
merged.sort((a, b) => a.id - b.id);
console.log(merged[0].name);
console.log(merged[0].age);
console.log(merged[1].name);
