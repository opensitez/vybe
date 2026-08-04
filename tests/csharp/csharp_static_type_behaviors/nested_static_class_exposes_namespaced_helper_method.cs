// vybe-test: csharp/csharp_static_type_behaviors/nested_static_class_exposes_namespaced_helper_method
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

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

class TextTools {
    public static class Parts {
        public static string Join(string a, string b) { return a + "/" + b; }
    }
}
__P((TextTools.Parts.Join("a", "b")).ToString());
__Check("a/b");
