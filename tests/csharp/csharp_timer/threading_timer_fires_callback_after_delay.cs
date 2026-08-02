// vybe-test: csharp/csharp_timer/threading_timer_fires_callback_after_delay
// origin: languages/csharp/tests/csharp/test_csharp_timer.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool fired=false;
using var t=new System.Threading.Timer(_=>{fired=true;},null,10,System.Threading.Timeout.Infinite);
System.Threading.Thread.Sleep(100);
__Check((fired).ToString(), "True");
