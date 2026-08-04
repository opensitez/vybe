// vybe-test: csharp/csharp_nested_classes/deeply_nested_class_visible_through_chain
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

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

class A{public class B{public class C{public int V=3;}}}
__P((new A.B.C().V).ToString());
__Check("3");
