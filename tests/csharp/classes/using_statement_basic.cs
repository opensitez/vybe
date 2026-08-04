// vybe-test: csharp/classes/using_statement_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

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

class Resource {
            public string name;
            public Resource(string n) { this.name = n; }
        }
        using (var r = new Resource("test")) {
            __P((r.name).ToString());
        }
__Check("test");
