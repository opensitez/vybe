// vybe-test: csharp/csharp_lock_monitor/lock_class_instance_field_gate
// origin: languages/csharp/tests/csharp/test_csharp_lock_monitor.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class SafeCounter {
    private object gate = new object();
    public int Value = 0;
    public void Add(int n) { lock (gate) { Value += n; } }
}
var sc = new SafeCounter();
sc.Add(3);
sc.Add(4);
__Check((sc.Value).ToString(), "7");
