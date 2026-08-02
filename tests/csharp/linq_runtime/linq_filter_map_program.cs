// vybe-test: csharp/linq_runtime/linq_filter_map_program
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var numbers = new List<int>();
numbers.Add(1); numbers.Add(2); numbers.Add(3); numbers.Add(4);
numbers.Add(5); numbers.Add(6); numbers.Add(7); numbers.Add(8);
var result = numbers.Where(n => n % 2 == 0).Select(n => n * n);
result.ForEach(x => Console.WriteLine(x));
