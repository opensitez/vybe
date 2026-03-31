// JS math utilities — compiled alongside VB in the same VM
// Both languages share the same host functions and globals

console.log("=== JS Module Loaded ===");

function factorial(n) {
    if (n <= 1) return 1;
    return n * factorial(n - 1);
}

console.log("factorial(5) = " + factorial(5));
console.log("Math.PI = " + Math.PI);
console.log("=== JS Done ===");
