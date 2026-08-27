// vybe-test: csharp/csharp_with_expression_records/with_nominal_two_inits
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var u=(new Theme{Name="dark",Ver=1}) with{Name="light",Ver=2}
;
__P((u.Name).ToString());
__P((u.Ver).ToString());
__Check("light\n2");

record Theme{public string Name{get;init;} public int Ver{get;init;}}

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
