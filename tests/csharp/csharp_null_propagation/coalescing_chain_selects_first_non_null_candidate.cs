// vybe-test: csharp/csharp_null_propagation/coalescing_chain_selects_first_non_null_candidate
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string first = null; string second = "B"; string third = "C"; __Check((first ?? second ?? third).ToString(), "B");
