// === Vybe JS VM Feature Test ===

// 1. Variables and types
let str = "hello";
let num = 42;
let bool_val = true;
let nothing = null;
console.log("1. Types:", str, num, bool_val, nothing);

// 2. Arithmetic
console.log("2. Math:", 2 + 3, 10 - 4, 3 * 7, 20 / 4, 17 % 5);

// 3. String concatenation
console.log("3. String:", "Hello" + " " + "World");

// 4. Comparison
console.log("4. Compare:", 5 > 3, 2 < 1, 10 === 10, "a" !== "b");

// 5. Recursive functions
function factorial(n) {
    if (n <= 1) {
        return 1;
    }
    return n * factorial(n - 1);
}
console.log("5. Factorial(6):", factorial(6));

// 6. Arrow functions
let square = (x) => x * x;
let add = (a, b) => a + b;
console.log("6. Arrow:", square(5), add(3, 4));

// 7. Objects
let person = { name: "Alice", age: 30 };
console.log("7. Object:", person.name, person.age);

// 8. Arrays
let nums = [10, 20, 30, 40, 50];
console.log("8. Array:", nums);

// 9. While loop
let fib1 = 0;
let fib2 = 1;
let count = 0;
while (count < 10) {
    let temp = fib2;
    fib2 = fib1 + fib2;
    fib1 = temp;
    count = count + 1;
}
console.log("9. Fib(10):", fib2);

// 10. For loop
let sum = 0;
for (let i = 1; i <= 100; i++) {
    sum = sum + i;
}
console.log("10. Sum 1-100:", sum);

// 11. Nested function calls
function compose(f, g) {
    return (x) => f(g(x));
}
let doubleSquare = compose(square, (x) => x * 2);
console.log("11. Compose:", doubleSquare(3));

// 12. Conditional (ternary)
let x = 10;
let result = x > 5 ? "big" : "small";
console.log("12. Ternary:", result);

// 13. Logical operators
console.log("13. Logic:", true && false, true || false, !false);

// 14. Switch
let day = 3;
switch (day) {
    case 1:
        console.log("14. Monday");
        break;
    case 3:
        console.log("14. Wednesday");
        break;
    default:
        console.log("14. Other");
}

// 15. Do-while
let n = 1;
do {
    n = n * 2;
} while (n < 100);
console.log("15. Do-while:", n);

console.log("=== All tests passed! ===");
