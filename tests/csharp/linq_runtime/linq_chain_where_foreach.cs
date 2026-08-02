// vybe-test: csharp/linq_runtime/linq_chain_where_foreach
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

var nums = new List<int>();
nums.Add(1); nums.Add(2); nums.Add(3); nums.Add(4); nums.Add(5);
nums.Where(x => x > 3).ForEach(x => Console.WriteLine(x));
