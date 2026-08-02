// vybe-test: csharp/csharp_yield_iterators_core/yield_return_in_if_branch_only_when_true
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<int> Pick(bool ok){if(ok){yield return 7;}yield return 0;}
__Check((string.Join(",",Pick(true))).ToString(), "7,0");
