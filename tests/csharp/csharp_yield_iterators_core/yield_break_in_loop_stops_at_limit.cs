// vybe-test: csharp/csharp_yield_iterators_core/yield_break_in_loop_stops_at_limit
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Take(int max){for(int i=0;i<10;i++){if(i>=max)yield break;yield return i;}}
Console.WriteLine(string.Join(",",Take(3)));
