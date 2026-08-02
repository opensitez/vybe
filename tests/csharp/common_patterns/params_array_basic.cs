// vybe-test: csharp/common_patterns/params_array_basic
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

class Logger {
    public static void Log(params string[] messages) {
        foreach (var m in messages) Console.WriteLine(m);
    }
}
Logger.Log("one", "two", "three");
