// vybe-test: csharp/classes/recursive_factorial
// origin: languages/csharp/tests/csharp/test_classes.rs

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

class MathUtils {
            public static int Fact(int n) {
                if (n <= 1) return 1;
                return Fact(n - 1) * n;
            }
        }
        __P((MathUtils.Fact(5)).ToString());
__Check("120");
