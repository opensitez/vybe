// vybe-test: csharp/advanced/class_type_local_decl
// origin: languages/csharp/tests/csharp/test_advanced.rs

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

class Foo {
            public string name;
            public Foo(string n) { this.name = n; }
        }
        Foo f = new Foo("hello");
        __P((f.name).ToString());
__Check("hello");
