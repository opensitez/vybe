// vybe-test: csharp/csharp_linq_set_ops/intersect_yields_elements_present_in_both_sequences
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

var result = new[]{1,2,3,4}.Intersect(new[]{2,4,6}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);
