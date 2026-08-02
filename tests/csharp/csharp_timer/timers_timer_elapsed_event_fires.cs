// vybe-test: csharp/csharp_timer/timers_timer_elapsed_event_fires
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool fired=false;
var t=new System.Timers.Timer(10){AutoReset=false};
t.Elapsed+=(_,__)=>fired=true;
t.Start();
System.Threading.Thread.Sleep(100);
__Check((fired).ToString(), "True");
