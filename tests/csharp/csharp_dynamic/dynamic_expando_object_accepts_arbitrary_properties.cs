// vybe-test: csharp/csharp_dynamic/dynamic_expando_object_accepts_arbitrary_properties
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

using static __Harness;

dynamic obj=new System.Dynamic.ExpandoObject();
obj.Name="Alice";
obj.Age=30;
__P((obj.Name).ToString());
__P((obj.Age).ToString());
__Check("Alice\n30");

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
