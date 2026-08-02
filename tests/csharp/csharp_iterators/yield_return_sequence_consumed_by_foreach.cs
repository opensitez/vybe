// vybe-test: csharp/csharp_iterators/yield_return_sequence_consumed_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1; yield return 2; yield return 3;
}
int sum = 0;
foreach(var n in Gen()) sum += n;
Console.WriteLine(sum);
