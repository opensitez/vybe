// vybe-test: csharp/csharp_exception_types/overflow_exception_thrown_in_checked_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try { checked { int x = int.MaxValue + 1; } }
catch(System.OverflowException) { result = "overflow"; }
__Check((result).ToString(), "overflow");
