// vybe-test: csharp/csharp_yield_iterators_core/yield_return_in_switch_case_arms
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.IEnumerable<string> Label(int n){switch(n){case 1:yield return "one";break;case 2:yield return "two";break;default:yield return "many";break;}}
__Check((string.Join("|",Label(2))).ToString(), "two");
