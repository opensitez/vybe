// vybe-test: csharp/classes/static_class_method_call
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathHelper {
            public static int Square(int x) { return x * x; }
            public static int Double(int x) { return x * 2; }
        }
        __Check((MathHelper.Square(5)).ToString(), "25");
        __Check((MathHelper.Double(7)).ToString(), "14");
