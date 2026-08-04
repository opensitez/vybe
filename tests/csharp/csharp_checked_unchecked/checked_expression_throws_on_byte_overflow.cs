// vybe-test: csharp/csharp_checked_unchecked/checked_expression_throws_on_byte_overflow
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

string r="";
try{byte b=checked((byte)256);}
catch(System.OverflowException){r="ov";}
__P((r).ToString());
__Check("ov");
