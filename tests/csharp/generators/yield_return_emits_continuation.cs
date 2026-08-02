// vybe-test: csharp/generators/yield_return_emits_continuation
// origin: languages/csharp/tests/csharp/test_generators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Program {
    public static IEnumerable<int> Count() {
        yield return 1;
        yield return 2;
        yield return 3;
    }
    public static void Run() {
        var g = Count();
        __Check((g).ToString(), "[continuation]");
    }
}
Program.Run();
