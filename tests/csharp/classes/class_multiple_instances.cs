// vybe-test: csharp/classes/class_multiple_instances
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
            int count;
            public Counter(int start) { this.count = start; }
            public void Inc() { this.count = this.count + 1; }
            public int Get() { return this.count; }
        }
        var a = new Counter(0);
        var b = new Counter(100);
        a.Inc(); a.Inc();
        b.Inc();
        __Check((a.Get()).ToString(), "2");
        __Check((b.Get()).ToString(), "101");
