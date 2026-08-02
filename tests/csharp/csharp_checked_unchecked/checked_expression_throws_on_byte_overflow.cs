// vybe-test: csharp/csharp_checked_unchecked/checked_expression_throws_on_byte_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="";
try{byte b=checked((byte)256);}
catch(System.OverflowException){r="ov";}
__Check((r).ToString(), "ov");
