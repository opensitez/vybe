function check(condition, message) {
    if (!condition) throw new Error(message);
}
const array = [];
const object = {};
let index = 0;
while (index < 1000) {
    array.push(index % 3 === 0 ? 'line\n"value' : index % 3 === 1 ? true : index);
    index++;
}
index = 0;
while (index < 500) {
    object['field' + index] = index % 2 === 0 ? 'value' : index;
    index++;
}
const arrayText = JSON.stringify(array);
const objectText = JSON.stringify(object);
function stringifyArray() { return JSON.stringify(array); }
function stringifyObject() { return JSON.stringify(object); }
function parseArray() { return JSON.parse(arrayText); }
function parseObject() { return JSON.parse(objectText); }
function validateArrayText(value) { check(value === arrayText, 'array text changed'); }
function validateObjectText(value) { check(value === objectText, 'object text changed'); }
function validateArray(value) {
    check(value.length === 1000 && value[0] === 'line\n"value' &&
        value[1] === true && value[998] === 998, 'array parse payload');
}
function validateObject(value) {
    check(Object.keys(value).length === 500 && value.field0 === 'value' &&
        value.field499 === 499, 'object parse payload');
}
function measure(label, operation, validate) {
    let warm = 0;
    while (warm < 10) { validate(operation()); warm++; }
    let count = 0;
    let result;
    const start = Date.now();
    while (count < 200) { result = operation(); count++; }
    const elapsed = Date.now() - start;
    validate(result);
    console.log(label + ': ' + elapsed + ' ms / 200 operations');
}
measure('JSON.stringify array 1000', stringifyArray, validateArrayText);
measure('JSON.stringify object 500', stringifyObject, validateObjectText);
measure('JSON.parse array 1000', parseArray, validateArray);
measure('JSON.parse object 500', parseObject, validateObject);
console.log('JSON workloads passed');
