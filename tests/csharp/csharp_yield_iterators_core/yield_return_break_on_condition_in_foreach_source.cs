// vybe-test: csharp/csharp_yield_iterators_core/yield_return_break_on_condition_in_foreach_source
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> TakeWhilePositive(int[] a){foreach(var n in a){if(n<0)yield break;yield return n;}}
Console.WriteLine(string.Join(",",TakeWhilePositive(new[]{2,4,-1,8})));
