// vybe-test: csharp/csharp_exception_types/invalid_cast_exception_thrown_by_explicit_reference_cast
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try { object o = "text"; int n = (int)o; }
catch(System.InvalidCastException) { result = "badcast"; }
__Check((result).ToString(), "badcast");
