// vybe-test: csharp/linq_compile/linq_aggregate_program
// origin: languages/csharp/tests/csharp/test_linq_compile.rs
// vybe-test-mode: compile

var words = new List<string>();
words.Add("Hello");
words.Add("World");
var first = words.First();
var last = words.Last();
Console.WriteLine(first);
Console.WriteLine(last);
