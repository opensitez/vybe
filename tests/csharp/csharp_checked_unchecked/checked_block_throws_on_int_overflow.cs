// vybe-test: csharp/csharp_checked_unchecked/checked_block_throws_on_int_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string r="ok";
try{checked{int x=int.MaxValue;x++;}}
catch(System.OverflowException){r="overflow";}
__Check((r).ToString(), "overflow");
