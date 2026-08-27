// vybe-test: csharp/csharp_struct_advanced/struct_with_operator_overload_usable_in_expressions
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

using static __Harness;

var r=new Fraction{Num=1,Den=2}
+new Fraction{Num=1,Den=3}
;
__P((r.Num).ToString());
__P((r.Den).ToString());
__Check("5\n6");

struct Fraction{
    public int Num,Den;
    public static Fraction operator+(Fraction a,Fraction b)=>
        new Fraction{Num=a.Num*b.Den+b.Num*a.Den,Den=a.Den*b.Den};
    public override string ToString()=>$"{Num}/{Den}";
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
