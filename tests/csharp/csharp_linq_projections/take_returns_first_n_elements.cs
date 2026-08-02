// vybe-test: csharp/csharp_linq_projections/take_returns_first_n_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{10,20,30,40}.Take(2);
foreach(var n in result) Console.WriteLine(n);
