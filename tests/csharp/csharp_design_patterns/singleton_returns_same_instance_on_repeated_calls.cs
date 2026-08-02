// vybe-test: csharp/csharp_design_patterns/singleton_returns_same_instance_on_repeated_calls
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Singleton{
    static Singleton _inst;
    public int Val;
    public static Singleton Instance=>_inst??=new Singleton();
}
Singleton.Instance.Val=42;
__Check((Singleton.Instance.Val).ToString(), "42");
