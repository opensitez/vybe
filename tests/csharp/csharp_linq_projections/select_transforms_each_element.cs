// vybe-test: csharp/csharp_linq_projections/select_transforms_each_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{1,2,3}.Select(x => x*x);
foreach(var n in result) Console.WriteLine(n);
