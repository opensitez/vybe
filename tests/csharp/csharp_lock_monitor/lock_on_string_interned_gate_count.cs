// vybe-test: csharp/csharp_lock_monitor/lock_on_string_interned_gate_count
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string gate = "sync-root";
int counter = 0;
lock (gate) { counter++; }
__Check((counter).ToString(), "1");
