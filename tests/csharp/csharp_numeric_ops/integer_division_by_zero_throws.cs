// vybe-test: csharp/csharp_numeric_ops/integer_division_by_zero_throws
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

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
try{int x=1/0;}
catch(System.DivideByZeroException){r="div0";}
__P((r).ToString());
__Check("div0");
