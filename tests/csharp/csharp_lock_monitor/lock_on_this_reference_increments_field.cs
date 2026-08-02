// vybe-test: csharp/csharp_lock_monitor/lock_on_this_reference_increments_field
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
    public int counter = 0;
    public void Inc() { lock (this) { counter++; } }
}
var box = new Box();
box.Inc();
__Check((box.counter).ToString(), "1");
