// vybe-test: csharp/csharp_nested_type_access/nested_access_two_nested_classes_same_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Pair{public class Left{public int V=1;} public class Right{public int V=2;}} __P((new Pair.Left().V).ToString()); __P((new Pair.Right().V).ToString());
__Check("1\n2");
