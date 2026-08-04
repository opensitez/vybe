// vybe-test: csharp/csharp_expression_bodied_members/expr_method_nested_class_delegates_to_outer
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

class Outer { public int Base => 10; public class Inner { Outer o; public Inner(Outer owner) { o = owner; } public int Boost() => o.Base + 5; } }
__P((new Outer.Inner(new Outer()).Boost()).ToString());
__Check("15");
