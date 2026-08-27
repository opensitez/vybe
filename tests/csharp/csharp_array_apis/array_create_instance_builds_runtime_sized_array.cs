// vybe-test: csharp/csharp_array_apis/array_create_instance_builds_runtime_sized_array
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var array = System.Array.CreateInstance(typeof(int), 3);
__P((array.Length).ToString());
__Check("3");

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
