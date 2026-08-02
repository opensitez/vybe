// vybe-test: csharp/csharp_decimal_semantics/decimal_increment_mutates_storage_in_place
// origin: languages/csharp/tests/csharp/test_csharp_decimal_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

decimal tally = 2.5m;
tally++;
__Check((tally).ToString(), "3.5");
