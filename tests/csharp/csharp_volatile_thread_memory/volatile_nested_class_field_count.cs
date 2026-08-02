// vybe-test: csharp/csharp_volatile_thread_memory/volatile_nested_class_field_count
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    public class Inner {
        public volatile int Value = 0;
    }
}
var inner = new Outer.Inner();
inner.Value = 13;
__Check((inner.Value).ToString(), "13");
