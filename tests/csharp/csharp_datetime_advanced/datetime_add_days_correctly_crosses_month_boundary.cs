// vybe-test: csharp/csharp_datetime_advanced/datetime_add_days_correctly_crosses_month_boundary
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

var d=new System.DateTime(2024,1,30).AddDays(3);
__P((d.Month).ToString());
__P((d.Day).ToString());
__Check("2\n2");

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
