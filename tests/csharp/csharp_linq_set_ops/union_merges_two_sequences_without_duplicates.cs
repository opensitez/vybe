// vybe-test: csharp/csharp_linq_set_ops/union_merges_two_sequences_without_duplicates
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

var result = new[]{1,2,3}.Union(new[]{3,4,5}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);
