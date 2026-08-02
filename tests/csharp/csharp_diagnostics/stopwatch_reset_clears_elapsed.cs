// vybe-test: csharp/csharp_diagnostics/stopwatch_reset_clears_elapsed
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sw=System.Diagnostics.Stopwatch.StartNew();
System.Threading.Thread.Sleep(5);
sw.Stop();
sw.Reset();
__Check((sw.Elapsed==System.TimeSpan.Zero).ToString(), "True");
