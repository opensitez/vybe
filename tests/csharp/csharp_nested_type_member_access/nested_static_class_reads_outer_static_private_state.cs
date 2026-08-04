// vybe-test: csharp/csharp_nested_type_member_access/nested_static_class_reads_outer_static_private_state
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

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

class Outer {
    static int tally = 3;
    static class Inner {
        public static int Read() { return tally; }
    }
    public static int Via() { return Inner.Read(); }
}
__P((Outer.Via()).ToString());
__Check("3");
