// vybe-test: csharp/csharp_yield_iterators_core/yield_return_with_local_state_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Running(){int s=0; for(int i=1;i<=3;i++){s+=i;yield return s;}}
Console.WriteLine(string.Join(",",Running()));
