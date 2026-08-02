// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_length_first_of_each_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

var r=new[]{"a","bb","c","dd","eee"}.DistinctBy(s=>s.Length);
foreach(var s in r) Console.WriteLine(s);
