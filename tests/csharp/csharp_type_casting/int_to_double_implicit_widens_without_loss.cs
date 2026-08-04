// vybe-test: csharp/csharp_type_casting/int_to_double_implicit_widens_without_loss
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

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

int i = 5; double d = i; __P((d).ToString());
__Check("5");
