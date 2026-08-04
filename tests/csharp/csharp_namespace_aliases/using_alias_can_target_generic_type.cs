// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_target_generic_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

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

using TextList = System.Collections.Generic.List<string>; var list = new TextList { "a", "b" }; __P((list.Count).ToString());
__Check("2");
