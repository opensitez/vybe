// vybe-test: csharp/advanced/class_type_local_decl
// origin: languages/csharp/tests/csharp/test_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Foo {
            public string name;
            public Foo(string n) { this.name = n; }
        }
        Foo f = new Foo("hello");
        __Check((f.name).ToString(), "hello");
