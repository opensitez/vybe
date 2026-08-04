// vybe-test: csharp/csharp_casting_patterns/direct_cast_succeeds_for_compatible_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

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

object o=3.14;
double d=(double)o;
__P((d).ToString());
__Check("3.14");
