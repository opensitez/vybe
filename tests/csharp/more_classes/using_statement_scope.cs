// vybe-test: csharp/more_classes/using_statement_scope
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

class Res {
            public int value;
            public Res(int v) { this.value = v; }
        }
        var total = 0;
        using (var r = new Res(42)) {
            total = r.value;
        }
        __P((total).ToString());
__Check("42");
