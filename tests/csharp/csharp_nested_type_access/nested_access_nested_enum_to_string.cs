// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_to_string
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Mode{public enum Kind{Alpha,Beta} public string Label()=>Kind.Beta.ToString();} __P((new Mode().Label()).ToString());
__Check("Beta");
