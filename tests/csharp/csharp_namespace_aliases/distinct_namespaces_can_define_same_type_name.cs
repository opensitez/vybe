// vybe-test: csharp/csharp_namespace_aliases/distinct_namespaces_can_define_same_type_name
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

namespace Left { public class Item { public string Name => "L"; } } namespace Right { public class Item { public string Name => "R"; } } __P((new Left.Item().Name).ToString()); __P((new Right.Item().Name).ToString());
__Check("L\nR");
