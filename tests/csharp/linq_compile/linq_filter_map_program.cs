// vybe-test: csharp/linq_compile/linq_filter_map_program
// origin: languages/csharp/tests/csharp/test_linq_compile.rs
// vybe-test-mode: compile

var numbers = new List<int>();
numbers.Add(1); numbers.Add(2); numbers.Add(3); numbers.Add(4);
numbers.Add(5); numbers.Add(6); numbers.Add(7); numbers.Add(8);
var evens = numbers.Where(n => n % 2 == 0);
var squared = evens.Select(n => n * n);
Console.WriteLine("Done");
