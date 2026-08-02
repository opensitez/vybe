// vybe-test: csharp/csharp_pattern_matching/not_pattern_inverts_any_sub_pattern
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o = "text"; __Check((o is not int).ToString(), "True");
