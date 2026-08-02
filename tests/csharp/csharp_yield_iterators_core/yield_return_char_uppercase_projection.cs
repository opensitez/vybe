// vybe-test: csharp/csharp_yield_iterators_core/yield_return_char_uppercase_projection
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<char> Lower(){yield return 'a';yield return 'b';}
__Check((string.Join(",",Lower().Select(c=>char.ToUpper(c)))).ToString(), "A,B");
