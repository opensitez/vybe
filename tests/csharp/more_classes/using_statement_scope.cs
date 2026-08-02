// vybe-test: csharp/more_classes/using_statement_scope
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Res {
            public int value;
            public Res(int v) { this.value = v; }
        }
        var total = 0;
        using (var r = new Res(42)) {
            total = r.value;
        }
        __Check((total).ToString(), "42");
