// vybe-test: csharp/csharp_method_group_delegates/method_group_converts_to_action_for_void_method
// origin: languages/csharp/tests/csharp/test_csharp_method_group_delegates.rs

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

int total = 0;
void Bump() { total++; }
System.Action bump = Bump;
bump();
__P((total).ToString());
__Check("1");
