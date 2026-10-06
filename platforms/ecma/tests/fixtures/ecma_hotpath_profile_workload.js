// Deterministic guest-level ECMA profiling workload.
// Phase clocks exclude compilation and initial VM/platform registration.
// Results are a current-build baseline, not proof of an application speedup.
const results = [];
const receiver = { offset: 10 };
const dense = [1, 2, 3];

function sum(a, b, c) {
    return this.offset + a + b + c;
}

function measure(name, iterations, workload, expected) {
    const start = Date.now();
    const checksum = workload(iterations);
    const elapsed = Date.now() - start;
    if (checksum !== expected) {
        throw new Error(name + ": incorrect checksum " + checksum);
    }
    results.push({ name: name, iterations: iterations, milliseconds: elapsed, checksum: checksum });
}

measure("apply_dense", 1000, function (count) {
    let checksum = 0;
    for (let i = 0; i < count; i++) {
        checksum += sum.apply(receiver, dense);
    }
    return checksum;
}, 16000);

let getterReads = 0;
const arrayLike = { length: 3, 0: 1, 2: 3 };
Object.defineProperty(arrayLike, "1", {
    get: function () {
        getterReads++;
        return 2;
    },
    configurable: true
});
measure("apply_array_like_getter", 100, function (count) {
    let checksum = 0;
    for (let i = 0; i < count; i++) {
        checksum += sum.apply(receiver, arrayLike);
    }
    return checksum;
}, 1600);
if (getterReads !== 100) {
    throw new Error("apply argument getter count: " + getterReads);
}

const source = { a: 1, b: 2, c: 3 };
measure("object_assign_keys_reflect", 1000, function (count) {
    let checksum = 0;
    for (let i = 0; i < count; i++) {
        const target = Object.assign({}, source);
        checksum += Object.keys(target).length + Reflect.get(target, "b") + target.c;
    }
    return checksum;
}, 8000);

measure("bind_create_and_call", 1000, function (count) {
    let checksum = 0;
    for (let i = 0; i < count; i++) {
        const bound = sum.bind(receiver, 1);
        checksum += bound(2, 3);
    }
    return checksum;
}, 16000);

const words = ["alpha", "beta", "gamma"];
measure("string_join_slice", 1000, function (count) {
    let checksum = 0;
    for (let i = 0; i < count; i++) {
        checksum += words.join(":").slice(1, 6).length;
    }
    return checksum;
}, 5000);

console.log(JSON.stringify({ phases: results, getterReads: getterReads }));
