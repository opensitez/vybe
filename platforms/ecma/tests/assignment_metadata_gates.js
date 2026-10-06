function check(condition, message) {
    if (!condition) throw new Error(message);
}
let calls = 0;
const child = {identity: 17};
const small = {first: 1, second: 2};
Object.defineProperty(small, 'unrelated', {
    get: function() { calls++; return 99; }
});
Object.assign(small, {first: child, second: 7});
check(small.first === child && small.second === 7, 'data writes with unrelated getter');
check(calls === 0, 'unrelated getter must not be invoked');
const readonly = {first: 1, second: 2};
Object.defineProperty(readonly, 'unrelated', {value: 99, writable: false});
Object.assign(readonly, {first: 7, second: child});
check(readonly.first === 7 && readonly.second === child && readonly.unrelated === 99,
    'unrelated readonly descriptor preserved');
const large = {};
let index = 0;
while (index < 1000) {
    large['field' + index] = index;
    index++;
}
Object.defineProperty(large, 'unrelated', {
    get: function() { calls++; return 99; }
});
Object.assign(large, {field0: child});
check(large.field0 === child && large.field999 === 999, 'small write into large target');
check(calls === 0, 'large target accessor stays unobserved');
const mixed = {first: 1};
Object.defineProperty(mixed, 'second', {value: 2, writable: false, enumerable: true});
let thrown = false;
try { Object.assign(mixed, {first: 7, second: 9, third: 11}); }
catch (error) { thrown = error instanceof TypeError; }
check(thrown && mixed.first === 7 && mixed.second === 2 && !('third' in mixed),
    'mixed readonly failure preserves ordered partial writes');
console.log('assignment metadata gates passed');
