// vybe-test: csharp/linq_runtime/linq_select_strings
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<string>();
list.Add("hello"); list.Add("world");
list.Select(s => s.ToUpper()).ForEach(s => Console.WriteLine(s));
