// vybe-test: csharp/csharp_static_classes/static_constructor_runs_once_before_first_member_access
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Singleton {
    public static int InitCount = 0;
    static Singleton() { InitCount++; }
    public static int Value = 42;
}
__Check((Singleton.Value).ToString(), "42");
__Check((Singleton.InitCount).ToString(), "1");
