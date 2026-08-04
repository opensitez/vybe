// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_static_method_group_to_action
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

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

static int hits = 0;
static void Hit() { hits++; }
System.Action strike = Hit;
strike(); strike();
__P((hits).ToString());
__Check("2");
