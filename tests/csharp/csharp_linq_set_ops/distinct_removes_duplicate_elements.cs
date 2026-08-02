// vybe-test: csharp/csharp_linq_set_ops/distinct_removes_duplicate_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

var result = new[]{1,2,2,3,1}.Distinct().OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);
