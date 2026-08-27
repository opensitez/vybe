// vybe-test: csharp/csharp_operators/user_defined_implicit_conversion_coerces_to_target_type
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

using static __Harness;

double length = new Inch { Value = 2.5 }
;
__P((length).ToString());
__Check("2.5");

struct Inch {
    public double Value;
    public static implicit operator double(Inch i) => i.Value;
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
