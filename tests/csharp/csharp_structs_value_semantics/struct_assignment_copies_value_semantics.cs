// vybe-test: csharp/csharp_structs_value_semantics/struct_assignment_copies_value_semantics
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

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

struct Counter { public int Value; } var left = new Counter { Value = 1 }; var right = left; right.Value = 9; __P((left.Value).ToString()); __P((right.Value).ToString());
__Check("1\n9");
