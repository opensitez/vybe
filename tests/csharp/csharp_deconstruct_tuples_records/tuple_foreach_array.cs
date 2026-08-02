// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_foreach_array
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

var pairs=new[]{(1,2),(3,4)}; int sum=0; foreach(var (x,y) in pairs) sum+=x+y; Console.WriteLine(sum);
