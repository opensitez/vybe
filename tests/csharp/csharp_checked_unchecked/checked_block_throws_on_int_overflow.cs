// vybe-test: csharp/csharp_checked_unchecked/checked_block_throws_on_int_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

string r="ok";
try{checked{int x=int.MaxValue;x++;}}
catch(System.OverflowException){r="overflow";}
__P((r).ToString());
__Check("overflow");
