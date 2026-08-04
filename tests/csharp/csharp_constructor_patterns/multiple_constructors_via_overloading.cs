// vybe-test: csharp/csharp_constructor_patterns/multiple_constructors_via_overloading
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

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

class Range{public int Lo,Hi;
    public Range():this(0,100){}
    public Range(int lo,int hi){Lo=lo;Hi=hi;}
}
var r1=new Range(); var r2=new Range(5,10);
__P((r1.Lo).ToString()); __P((r2.Hi).ToString());
__Check("0\n10");
