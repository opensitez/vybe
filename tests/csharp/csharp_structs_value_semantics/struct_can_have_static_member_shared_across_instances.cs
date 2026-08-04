// vybe-test: csharp/csharp_structs_value_semantics/struct_can_have_static_member_shared_across_instances
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

struct Token { public static int Count; public Token(int _) { Count++; } } new Token(1); new Token(2); __P((Token.Count).ToString());
__Check("2");
