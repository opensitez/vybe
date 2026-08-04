// vybe-test: csharp/type_gettype_dynamic/get_type_returns_null_for_an_unknown_name
// origin: languages/csharp/tests/csharp/test_type_gettype_dynamic.rs

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

class Program {
    static void Main() {
        System.__P((System.Type.GetType("NoSuchTypeAnywhere") == null).ToString());
    }
}
__Check("True");
