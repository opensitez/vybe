// vybe-test: csharp/csharp_lock_monitor/lock_switch_case_increments_matching
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
int code = 2;
lock (gate) {
    switch (code) {
        case 1: counter = 10; break;
        case 2: counter = 20; break;
        default: counter = 0; break;
    }
}
__Check((counter).ToString(), "20");
