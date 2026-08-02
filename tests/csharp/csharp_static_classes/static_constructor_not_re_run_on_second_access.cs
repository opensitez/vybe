// vybe-test: csharp/csharp_static_classes/static_constructor_not_re_run_on_second_access
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Registry {
    public static int Boot = 0;
    static Registry() { Boot++; }
    public static void Touch() { }
}
Registry.Touch();
Registry.Touch();
__Check((Registry.Boot).ToString(), "1");
