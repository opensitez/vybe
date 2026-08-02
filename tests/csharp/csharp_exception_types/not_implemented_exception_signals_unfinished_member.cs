// vybe-test: csharp/csharp_exception_types/not_implemented_exception_signals_unfinished_member
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try { throw new System.NotImplementedException(); }
catch(System.NotImplementedException) { result = "ni"; }
__Check((result).ToString(), "ni");
