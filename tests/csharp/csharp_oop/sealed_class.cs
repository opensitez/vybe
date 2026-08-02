// vybe-test: csharp/csharp_oop/sealed_class
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

sealed class Singleton {
    public int Value = 42;
}
var s = new Singleton();
__Check((s.Value).ToString(), "42");
