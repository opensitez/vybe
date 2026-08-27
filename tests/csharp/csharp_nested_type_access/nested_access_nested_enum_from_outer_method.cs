// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_from_outer_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Job().Current()).ToString());
__Check("Busy");

class Job{public enum State{Idle,Busy} public State Current()=>State.Busy;}

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
