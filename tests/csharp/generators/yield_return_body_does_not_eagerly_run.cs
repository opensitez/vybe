// vybe-test: csharp/generators/yield_return_body_does_not_eagerly_run
// origin: languages/csharp/tests/csharp/test_generators.rs

class Program {
    public static IEnumerable<int> Loud() {
        Console.WriteLine("bad: body ran without resume");
        yield return 1;
    }
    public static void Run() {
        var _ = Loud();
        Console.WriteLine("ok");
    }
}
Program.Run();
