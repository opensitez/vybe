// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_preserves_first_occurrence_foreach
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

var r=new[]{2,1,2,3,1}.Distinct();
foreach(var n in r) Console.WriteLine(n);
