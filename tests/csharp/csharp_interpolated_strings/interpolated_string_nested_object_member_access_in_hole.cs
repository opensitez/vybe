// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_nested_object_member_access_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

using static __Harness;

var pair = new Pair { A = 2, B = 3 }
;
__P(($"{pair.A + pair.B}").ToString());
__Check("5");

class Pair { public int A; public int B; }

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
