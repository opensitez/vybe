// vybe-test: csharp/csharp_linq_set_ops/except_yields_elements_in_first_not_in_second
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

var result = new[]{1,2,3,4}.Except(new[]{2,4}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);
