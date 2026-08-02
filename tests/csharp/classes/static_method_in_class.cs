// vybe-test: csharp/classes/static_method_in_class
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtils {
            public static int Add(int a, int b) { return a + b; }
        }
        __Check((MathUtils.Add(3, 4)).ToString(), "7");
