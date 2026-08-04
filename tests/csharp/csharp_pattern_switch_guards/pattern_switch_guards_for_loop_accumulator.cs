// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

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

// pattern_switch_guards
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } __P((sum == 3).ToString());
__Check("True");
