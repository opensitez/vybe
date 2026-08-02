// vybe-test: csharp/csharp_null_handling/null_coalescing_chained_selects_first_non_null
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a=null, b=null, c="found";
__Check((a ?? b ?? c).ToString(), "found");
