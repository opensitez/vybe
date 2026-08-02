// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_discards_unused_value
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (name, _) = ("Ada", 99);
__Check((name).ToString(), "Ada");
