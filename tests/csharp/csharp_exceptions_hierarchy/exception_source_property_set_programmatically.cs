// vybe-test: csharp/csharp_exceptions_hierarchy/exception_source_property_set_programmatically
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

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

var ex=new System.Exception("e");
ex.Source="MyModule";
__P((ex.Source).ToString());
__Check("MyModule");
