// vybe-test: csharp/csharp_yield_iterators_core/yield_return_filter_with_continue_skip
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Evens(int max){for(int i=0;i<=max;i++){if(i%2!=0)continue;yield return i;}}
Console.WriteLine(string.Join(",",Evens(6)));
