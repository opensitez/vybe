// vybe-test: csharp/more_classes/return_object
// origin: languages/csharp/tests/csharp/test_more_classes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Result {
            public int value;
            public bool ok;
            public Result(int v, bool o) { this.value = v; this.ok = o; }
        }
        var r = new Result(42, true);
        __P((r.value).ToString());
        __P((r.ok).ToString());
__Check("42\nTrue");
