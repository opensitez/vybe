// vybe-test: csharp/linq_runtime/linq_select_double
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(1); list.Add(2); list.Add(3);
list.Select(x => x * 2).ForEach(x => Console.WriteLine(x));
