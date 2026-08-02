// vybe-test: csharp/csharp_numeric_ops/integer_division_by_zero_throws
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{int x=1/0;}
catch(System.DivideByZeroException){r="div0";}
__Check((r).ToString(), "div0");
