// vybe-test: csharp/csharp_nested_control_flow/foreach_continue_skips_even_numbers_only
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

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

int sum = 0;
foreach (var value in new[] { 1, 2, 3, 4, 5 }) {
    if (value % 2 == 0) continue;
    sum += value;
}
__P((sum).ToString());
__Check("9");
