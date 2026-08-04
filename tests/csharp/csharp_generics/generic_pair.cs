// vybe-test: csharp/csharp_generics/generic_pair
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

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

class Pair<T1, T2> {
    public T1 First;
    public T2 Second;
    public Pair(T1 a, T2 b) { First = a; Second = b; }
}
var p = new Pair<string, int>("age", 30);
__P((p.First).ToString());
__P((p.Second).ToString());
__Check("age\n30");
