// vybe-test: csharp/csharp_structs_value_semantics/struct_can_override_to_string_for_custom_output
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

struct Token { public int Value; public override string ToString() { return "T:" + Value; } } __P((new Token { Value = 7 }).ToString());
__Check("T:7");
