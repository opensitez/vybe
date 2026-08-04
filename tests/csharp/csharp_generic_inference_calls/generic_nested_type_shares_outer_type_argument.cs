// vybe-test: csharp/csharp_generic_inference_calls/generic_nested_type_shares_outer_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

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

class Outer<T> {
    public class Inner {
        public T Value;
    }
    public Inner Build(T value) {
        return new Inner { Value = value };
    }
}
var built = new Outer<string>().Build("nested");
__P((built.Value).ToString());
__Check("nested");
