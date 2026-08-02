// vybe-test: csharp/csharp_static_classes/static_method_can_call_other_static_methods_in_same_class
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Calc {
    static int Add(int a, int b) => a+b;
    public static int Sum3(int a, int b, int c) => Add(Add(a,b),c);
}
__Check((Calc.Sum3(1,2,3)).ToString(), "6");
