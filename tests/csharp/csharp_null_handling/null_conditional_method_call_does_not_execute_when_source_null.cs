// vybe-test: csharp/csharp_null_handling/null_conditional_method_call_does_not_execute_when_source_null
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s = null;
int count = 0;
s?.ToUpper();
__Check((count).ToString(), "0");
