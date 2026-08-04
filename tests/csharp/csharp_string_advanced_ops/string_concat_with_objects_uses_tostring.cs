// vybe-test: csharp/csharp_string_advanced_ops/string_concat_with_objects_uses_tostring
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

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

__P((string.Concat("val=",42," ok=",true)).ToString());
__Check("val=42 ok=True");
