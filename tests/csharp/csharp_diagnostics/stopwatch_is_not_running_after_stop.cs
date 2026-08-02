// vybe-test: csharp/csharp_diagnostics/stopwatch_is_not_running_after_stop
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sw=System.Diagnostics.Stopwatch.StartNew();
sw.Stop();
__Check((sw.IsRunning).ToString(), "False");
