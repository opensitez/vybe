// vybe-test: csharp/multi_value/tuple_return_and_deconstruct
// origin: languages/csharp/tests/csharp/test_multi_value.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Program {
    public static (int, int) Swap(int a, int b) {
        return (b, a);
    }
    public static void Run() {
        var (x, y) = Swap(1, 2);
        __Check((x).ToString(), "2");
        __Check((y).ToString(), "1");
    }
}
Program.Run();
