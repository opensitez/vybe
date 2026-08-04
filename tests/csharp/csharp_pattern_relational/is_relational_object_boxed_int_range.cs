// vybe-test: csharp/csharp_pattern_relational/is_relational_object_boxed_int_range
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

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

object o=7; __P((o is int n and >=7 and <=7).ToString());
__Check("True");
