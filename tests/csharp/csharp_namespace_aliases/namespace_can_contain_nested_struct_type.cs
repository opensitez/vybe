// vybe-test: csharp/csharp_namespace_aliases/namespace_can_contain_nested_struct_type
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

namespace Demo { public struct Point { public int X; public int Y; } } var point = new Demo.Point { X = 2, Y = 5 }; __P((point.X + point.Y).ToString());
__Check("7");
