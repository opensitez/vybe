// vybe-test: csharp/classes/property_chain
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner { public int value; public Inner(int v) { this.value = v; } }
        class Outer { public Inner inner; public Outer(int v) { this.inner = new Inner(v); } }
        var o = new Outer(42);
        __Check((o.inner.value).ToString(), "42");
