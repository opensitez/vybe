// vybe-test: csharp/csharp_string_advanced_ops/string_join_with_ienumerable_source
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums=Enumerable.Range(1,5);
__Check((string.Join("-",nums)).ToString(), "1-2-3-4-5");
