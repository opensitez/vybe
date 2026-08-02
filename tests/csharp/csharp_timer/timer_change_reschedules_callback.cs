// vybe-test: csharp/csharp_timer/timer_change_reschedules_callback
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count=0;
using var t=new System.Threading.Timer(_=>System.Threading.Interlocked.Increment(ref count),null,10,10);
System.Threading.Thread.Sleep(100);
__Check((count>0).ToString(), "True");
