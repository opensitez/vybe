// vybe-test: csharp/csharp_string_join_split/join_ienumerable_range
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Join("+",Enumerable.Range(1,4))).ToString(), "1+2+3+4");
