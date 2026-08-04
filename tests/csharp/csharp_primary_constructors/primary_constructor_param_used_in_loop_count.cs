// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_loop_count
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

class Repeat(int times) { public int Run() { int t = 0; for (int i = 0; i < times; i++) t++; return t; } }
__P((new Repeat(4).Run()).ToString());
__Check("4");
