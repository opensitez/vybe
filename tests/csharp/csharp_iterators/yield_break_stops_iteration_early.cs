// vybe-test: csharp/csharp_iterators/yield_break_stops_iteration_early
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1;
    yield break;
    yield return 2;
}
int count = 0;
foreach(var _ in Gen()) count++;
Console.WriteLine(count);
