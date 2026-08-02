// vybe-test: csharp/csharp_null_handling/null_conditional_member_access_returns_null_when_source_null
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
__Check((s?.Length == null).ToString(), "True");
