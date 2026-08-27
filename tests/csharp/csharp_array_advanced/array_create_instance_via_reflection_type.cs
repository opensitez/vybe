// vybe-test: csharp/csharp_array_advanced/array_create_instance_via_reflection_type
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

using static __Harness;

var arr=(int[])System.Array.CreateInstance(typeof(int),5);
arr[3]=99;
__P((arr[3]).ToString());
__Check("99");

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
