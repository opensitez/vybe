// vybe-test: csharp/advanced/null_conditional_access
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Foo { public string name; public Foo(string n) { this.name = n; } }
        var f = new Foo("test");
        __Check((f?.name).ToString(), "test");
