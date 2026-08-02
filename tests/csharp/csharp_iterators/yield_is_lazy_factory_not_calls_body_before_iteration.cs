// vybe-test: csharp/csharp_iterators/yield_is_lazy_factory_not_calls_body_before_iteration
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

int calls=0;
System.Collections.Generic.IEnumerable<int> Lazy() {
    calls++;
    yield return 1;
}
Console.WriteLine(calls);
var seq = Lazy();
Console.WriteLine(calls);
foreach(var _ in seq) {}
Console.WriteLine(calls);
