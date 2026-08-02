// vybe-test: csharp/csharp_diagnostics/stopwatch_restart_starts_from_zero
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sw=new System.Diagnostics.Stopwatch();
sw.Start();
System.Threading.Thread.Sleep(5);
sw.Restart();
__Check((sw.IsRunning).ToString(), "True");
