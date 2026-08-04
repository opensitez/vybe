// vybe-test: csharp/common_patterns/gcd_euclidean
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

class Algorithms {
    public static int GCD(int a, int b) {
        while (b != 0) { int t = b; b = a % b; a = t; }
        return a;
    }
}
__P((Algorithms.GCD(48, 18)).ToString());
__P((Algorithms.GCD(100, 75)).ToString());
__Check("6\n25");
