// vybe-test: csharp/csharp_scope_variables/const_local_cannot_be_reassigned_but_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

const int MAX = 100;
__Check((MAX).ToString(), "100");
