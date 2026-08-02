// vybe-test: csharp/csharp_exception_types/divide_by_zero_exception_thrown_by_integer_division
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
try { int x = 10 / 0; }
catch(System.DivideByZeroException e) { result = "dbz"; }
__Check((result).ToString(), "dbz");
