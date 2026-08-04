// vybe-test: csharp/csharp_nested_partial_types/nested_enum_values_are_accessible_through_outer_type
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Job {
    public enum State { Pending, Running, Done }
}
__P((Job.State.Pending).ToString());
__P(((int)Job.State.Done).ToString());
__Check("Pending\n2");
