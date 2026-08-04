// vybe-test: csharp/csharp_nested_classes/nested_class_can_access_outer_private_members
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

class Outer{
    static int secret=42;
    public class Inner{public int Get()=>secret;}
}
__P((new Outer.Inner().Get()).ToString());
__Check("42");
