// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_in_foreach_reads_current_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

int[] data={3,6,9}; int sum=0; foreach(ref readonly int n in data){sum+=n;} Console.WriteLine(sum);
