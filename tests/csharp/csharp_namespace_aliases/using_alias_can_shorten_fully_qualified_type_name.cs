// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_shorten_fully_qualified_type_name
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

using Thing = Demo.Tools.Box; namespace Demo.Tools { public class Box { public int Value = 7; } } __P((new Thing().Value).ToString());
__Check("7");
