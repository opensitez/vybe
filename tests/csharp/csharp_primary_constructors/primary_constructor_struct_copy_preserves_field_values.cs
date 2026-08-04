// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copy_preserves_field_values
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

struct Pair(int a, int b) { public int A = a; public int B = b; }
var p = new Pair(2, 3);
var q = p;
__P((q.A + q.B).ToString());
__Check("5");
