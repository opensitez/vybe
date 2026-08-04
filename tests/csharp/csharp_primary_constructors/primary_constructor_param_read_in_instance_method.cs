// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_read_in_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Counter(int start) {
    int current = start;
    public int Next() => ++current;
    public int Value => current;
}
var c = new Counter(10);
c.Next(); c.Next();
__P((c.Value).ToString());
__Check("12");
