// vybe-test: csharp/csharp_primary_constructors/primary_constructor_method_mutates_field_not_param
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

class Acc(int start) { int total = start; public void Add(int n) { total += n; } public int Value => total; }
var a = new Acc(1);
a.Add(4);
__P((a.Value).ToString());
__Check("5");
