// vybe-test: csharp/csharp_delegate_types/multicast_delegate_invokes_both_handlers_in_order
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

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

System.Action log = () => __P(("a").ToString());
log += () => __P(("b").ToString());
log();
__Check("a\nb");
