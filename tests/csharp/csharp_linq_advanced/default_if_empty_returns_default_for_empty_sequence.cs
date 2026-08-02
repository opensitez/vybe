// vybe-test: csharp/csharp_linq_advanced/default_if_empty_returns_default_for_empty_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=System.Array.Empty<int>().DefaultIfEmpty(99);
__Check((result.First()).ToString(), "99");
