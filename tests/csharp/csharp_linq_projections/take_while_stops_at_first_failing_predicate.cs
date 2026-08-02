// vybe-test: csharp/csharp_linq_projections/take_while_stops_at_first_failing_predicate
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

var result = new[]{1,3,5,4,7}.TakeWhile(x => x%2!=0);
foreach(var n in result) Console.WriteLine(n);
