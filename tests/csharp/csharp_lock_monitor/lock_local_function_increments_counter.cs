// vybe-test: csharp/csharp_lock_monitor/lock_local_function_increments_counter
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object gate = new object();
int counter = 0;
void Bump() { lock (gate) { counter++; } }
Bump();
Bump();
__Check((counter).ToString(), "2");
