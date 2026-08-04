// vybe-test: csharp/csharp_using_disposal/lock_statement_can_mutate_shared_local_state
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

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

object gate = new object(); int count = 0; lock (gate) { count += 3; } __P((count).ToString());
__Check("3");
