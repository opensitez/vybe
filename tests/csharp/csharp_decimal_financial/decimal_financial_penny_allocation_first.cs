// vybe-test: csharp/csharp_decimal_financial/decimal_financial_penny_allocation_first
// origin: languages/csharp/tests/csharp/test_csharp_decimal_financial.rs

using static __Harness;

decimal total=0.10m;
int parts=3;
decimal share=decimal.Truncate(total/parts*100m)/100m;
__P((share).ToString());
__Check("0.03");

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
