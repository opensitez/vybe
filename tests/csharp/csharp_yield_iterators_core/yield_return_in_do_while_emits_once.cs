// vybe-test: csharp/csharp_yield_iterators_core/yield_return_in_do_while_emits_once
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<int> Once(){int n=0;do{yield return n;n++;}while(n<1);}
Console.WriteLine(string.Join(",",Once()));
