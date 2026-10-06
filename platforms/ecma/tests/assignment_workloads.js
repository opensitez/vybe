// Guest-level workload: setup and warmup excluded from each timed section.
// Compare new-property writes with overwrites without changing host/compiler ABI.
function check(condition, message) {
    if (!condition) throw new Error(message);
}
function measure(source, size, fresh) {
    const target = {};
    Object.assign(target, source);
    let warm = 0;
    while (warm < 100) {
        Object.assign(fresh ? {} : target, source);
        warm++;
    }
    const iterations = 2000;
    let result = target;
    let index = 0;
    const start = Date.now();
    while (index < iterations) {
        result = Object.assign(fresh ? {} : target, source);
        index++;
    }
    const elapsed = Date.now() - start;
    check(result.field0 === 17, 'first property copied');
    check(Object.keys(result).length === size, 'property count preserved');
    check(fresh || result === target, 'overwrite preserves target identity');
    source.field0 = 29;
    Object.assign(result, source);
    check(result.field0 === 29, 'later assignment observes source mutation');
    source.field0 = 17;
    console.log('assignment size=' + size + ' fresh=' + fresh +
        ' iterations=' + iterations + ' elapsed_ms=' + elapsed);
}
const one = {field0: 17};
const four = {field0: 17, field1: 2, field2: 3, field3: 4};
const sixteen = {field0: 17, field1: 2, field2: 3, field3: 4,
    field4: 5, field5: 6, field6: 7, field7: 8,
    field8: 9, field9: 10, field10: 11, field11: 12,
    field12: 13, field13: 14, field14: 15, field15: 16};
measure(one, 1, false);
measure(one, 1, true);
measure(four, 4, false);
measure(four, 4, true);
measure(sixteen, 16, false);
measure(sixteen, 16, true);
console.log('assignment guest workload checks passed');
