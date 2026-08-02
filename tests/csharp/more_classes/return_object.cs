// vybe-test: csharp/more_classes/return_object
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Result {
            public int value;
            public bool ok;
            public Result(int v, bool o) { this.value = v; this.ok = o; }
        }
        var r = new Result(42, true);
        __Check((r.value).ToString(), "42");
        __Check((r.ok).ToString(), "True");
