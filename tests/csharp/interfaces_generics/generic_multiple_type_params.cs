// vybe-test: csharp/interfaces_generics/generic_multiple_type_params
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

class Pair<TFirst, TSecond> {
    public TFirst First;
    public TSecond Second;
    public Pair(TFirst f, TSecond s) { First = f; Second = s; }
    public override string ToString() { return First + ":" + Second; }
}
var p = new Pair<string, int>("age", 30);
__P((p).ToString());
__Check("age:30");
