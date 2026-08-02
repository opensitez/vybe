// vybe-test: csharp/classes/class_field_access
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box {
            public int value;
            public Box(int v) { this.value = v; }
        }
        var b = new Box(42);
        __Check((b.value).ToString(), "42");
