// vybe-test: csharp/csharp_iterators/yield_in_loop_produces_computed_sequence
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

System.Collections.Generic.IEnumerable<int> Range(int n) {
    for(int i=0; i<n; i++) yield return i;
}
int sum=0;
foreach(var x in Range(5)) sum+=x;
Console.WriteLine(sum);
