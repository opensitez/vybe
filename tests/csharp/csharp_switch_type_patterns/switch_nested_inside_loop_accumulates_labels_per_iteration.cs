// vybe-test: csharp/csharp_switch_type_patterns/switch_nested_inside_loop_accumulates_labels_per_iteration
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

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

string trace = "";
for (int i = 0; i < 3; i++) {
    trace += i switch { 0 => "a", 1 => "b", _ => "c" };
}
__P((trace).ToString());
__Check("abc");
