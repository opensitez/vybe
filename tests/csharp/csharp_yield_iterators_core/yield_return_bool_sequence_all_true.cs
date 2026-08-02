// vybe-test: csharp/csharp_yield_iterators_core/yield_return_bool_sequence_all_true
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<bool> Flags(){yield return true;yield return true;}
__Check((Flags().All(x=>x)).ToString(), "True");
