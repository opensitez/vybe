// vybe-test: csharp/csharp_yield_iterators_core/yield_return_string_chars_sequence
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

System.Collections.Generic.IEnumerable<char> Letters(string s){foreach(char c in s)yield return c;}
Console.WriteLine(string.Join("",Letters("ab")));
