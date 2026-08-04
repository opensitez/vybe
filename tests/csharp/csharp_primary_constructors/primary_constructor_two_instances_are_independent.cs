// vybe-test: csharp/csharp_primary_constructors/primary_constructor_two_instances_are_independent
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

class Slot(int id) { public int Id => id; }
var a = new Slot(1);
var b = new Slot(2);
__P((a.Id).ToString()); __P((b.Id).ToString());
__Check("1\n2");
