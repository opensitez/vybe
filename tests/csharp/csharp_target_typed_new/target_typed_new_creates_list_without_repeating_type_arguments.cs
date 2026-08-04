// vybe-test: csharp/csharp_target_typed_new/target_typed_new_creates_list_without_repeating_type_arguments
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new.rs

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

System.Collections.Generic.List<int> values = new();
values.Add(7);
__P((values[0]).ToString());
__Check("7");
