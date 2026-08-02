// vybe-test: csharp/csharp_static_classes/static_class_method_callable_without_instance
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class MathHelper { public static int Square(int n) => n*n; }
__Check((MathHelper.Square(5)).ToString(), "25");
