// vybe-test: csharp/csharp_patterns/accumulator_pattern
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

var items = new[] { 1, 2, 3, 4, 5 };
int sum = 0;
int product = 1;
foreach (var x in items) {
    sum += x;
    product *= x;
}
Console.WriteLine(sum);
Console.WriteLine(product);
