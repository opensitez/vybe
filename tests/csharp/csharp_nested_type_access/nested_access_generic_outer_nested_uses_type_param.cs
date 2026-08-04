// vybe-test: csharp/csharp_nested_type_access/nested_access_generic_outer_nested_uses_type_param
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

class Box<T>{public class Holder{public T Value;}} var h=new Box<int>.Holder(); h.Value=15; __P((h.Value).ToString());
__Check("15");
