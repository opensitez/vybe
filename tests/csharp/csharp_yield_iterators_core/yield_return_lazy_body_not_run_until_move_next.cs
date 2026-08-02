// vybe-test: csharp/csharp_yield_iterators_core/yield_return_lazy_body_not_run_until_move_next
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

int calls=0; System.Collections.Generic.IEnumerable<int> Lazy(){calls++;yield return 1;}
var seq=Lazy(); Console.WriteLine(calls); foreach(var _ in seq){} Console.WriteLine(calls);
