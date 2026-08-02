// vybe-test: csharp/classes/recursive_factorial
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtils {
            public static int Fact(int n) {
                if (n <= 1) return 1;
                return Fact(n - 1) * n;
            }
        }
        __Check((MathUtils.Fact(5)).ToString(), "120");
