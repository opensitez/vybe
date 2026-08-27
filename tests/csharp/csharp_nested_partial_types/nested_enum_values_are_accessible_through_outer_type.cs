// vybe-test: csharp/csharp_nested_partial_types/nested_enum_values_are_accessible_through_outer_type
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

__P((Job.State.Pending).ToString());
__P(((int)Job.State.Done).ToString());
__Check("Pending\n2");

class Job {
    public enum State { Pending, Running, Done }
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
