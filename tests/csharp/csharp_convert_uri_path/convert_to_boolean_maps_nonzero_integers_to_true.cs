// vybe-test: csharp/csharp_convert_uri_path/convert_to_boolean_maps_nonzero_integers_to_true
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

using static __Harness;

// convert_uri_path
__P((System.Convert.ToBoolean(1)).ToString());
__Check("True");

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
