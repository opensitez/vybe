// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_class_equality
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

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

class Tag { public string Name; public static bool operator ==(Tag a, Tag b) => a.Name == b.Name; public static bool operator !=(Tag a, Tag b) => !(a == b); }
__P((new Tag { Name = "x" } == new Tag { Name = "x" }).ToString()); __P((new Tag { Name = "a" } != new Tag { Name = "b" }).ToString());
__Check("True\nTrue");
