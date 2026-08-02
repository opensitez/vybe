// vybe-test: csharp/csharp_linq_set_ops/zip_pairs_elements_from_two_sequences_by_position
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

var result = new[]{1,2,3}.Zip(new[]{10,20,30}, (a,b) => a*b);
foreach(var x in result) Console.WriteLine(x);
