// vybe-test: csharp/csharp_constructor_patterns/multiple_constructors_via_overloading
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

using static __Harness;

var r1=new Range();
var r2=new Range(5,10);
__P((r1.Lo).ToString());
__P((r2.Hi).ToString());
__Check("0\n10");

class Range{public int Lo,Hi;
    public Range():this(0,100){}
    public Range(int lo,int hi){Lo=lo;Hi=hi;}
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
