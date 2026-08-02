// vybe-test: csharp/csharp_iterators/multiple_foreach_iterations_restart_the_iterator
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

System.Collections.Generic.IEnumerable<int> Three() {
    yield return 1; yield return 2; yield return 3;
}
int total=0;
foreach(var x in Three()) total+=x;
foreach(var x in Three()) total+=x;
Console.WriteLine(total);
