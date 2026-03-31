// Hello World - basic JS test
console.log("Hello from JavaScript!");
console.log("Running on the Vybe bytecode VM");

// Variables and arithmetic
let x = 10;
let y = 20;
console.log("x + y =", x + y);

// Function
function greet(name) {
    return "Hello, " + name + "!";
}
console.log(greet("World"));

// Conditional
let n = 42;
if (n > 40) {
    console.log("n is greater than 40");
} else {
    console.log("n is 40 or less");
}

// Loop
let sum = 0;
for (let i = 1; i <= 10; i++) {
    sum = sum + i;
}
console.log("Sum 1..10 =", sum);

// Arrow function
let double = (x) => x * 2;
console.log("double(21) =", double(21));

// Array
let arr = [1, 2, 3, 4, 5];
console.log("Array:", arr);

// Object
let obj = { name: "Vybe", version: 1 };
console.log("Object name:", obj.name);

// Template literal
let lang = "JavaScript";
console.log("Running " + lang + " on Vybe!");
