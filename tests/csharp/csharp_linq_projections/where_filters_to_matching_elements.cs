// vybe-test: csharp/csharp_linq_projections/where_filters_to_matching_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{1,2,3,4,5}.Where(x => x%2==0);
foreach(var n in result) Console.WriteLine(n);
