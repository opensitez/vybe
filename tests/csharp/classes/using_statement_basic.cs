// vybe-test: csharp/classes/using_statement_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Resource {
            public string name;
            public Resource(string n) { this.name = n; }
        }
        using (var r = new Resource("test")) {
            __Check((r.name).ToString(), "test");
        }
