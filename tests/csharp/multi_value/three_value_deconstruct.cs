// vybe-test: csharp/multi_value/three_value_deconstruct
// origin: languages/csharp/tests/csharp/test_multi_value.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Program {
    public static (int, int, int) Rgb() {
        return (10, 20, 30);
    }
    public static void Run() {
        var (r, g, b) = Rgb();
        __Check((r).ToString(), "10");
        __Check((g).ToString(), "20");
        __Check((b).ToString(), "30");
    }
}
Program.Run();
