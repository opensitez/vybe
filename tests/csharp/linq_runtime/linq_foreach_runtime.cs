// vybe-test: csharp/linq_runtime/linq_foreach_runtime
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var list = new List<int>();
list.Add(10); list.Add(20); list.Add(30);
list.ForEach(x => Console.WriteLine(x));
