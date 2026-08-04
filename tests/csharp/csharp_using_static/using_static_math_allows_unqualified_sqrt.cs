// vybe-test: csharp/csharp_using_static/using_static_math_allows_unqualified_sqrt
// origin: languages/csharp/tests/csharp/test_csharp_using_static.rs

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

using static System.Math;
__P((Sqrt(16)).ToString());
__Check("4");
