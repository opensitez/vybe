// vybe-test: csharp/csharp_yield_iterators_core/yield_return_enumerable_of_enumerable_flatten_manual
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<System.Collections.Generic.IEnumerable<int>> Batches(){yield return new[]{1,2};yield return new[]{3};}
var flat=new System.Collections.Generic.List<int>(); foreach(var batch in Batches()) foreach(var n in batch) flat.Add(n); Console.WriteLine(string.Join(",",flat));
