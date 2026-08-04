// vybe-test: csharp/csharp_abstract_sealed/sealed_class_cannot_be_used_as_base_detected_at_compile_time_but_runtime_ok
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

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

sealed class Final { public int Value = 7; }
var f = new Final();
__P((f.Value).ToString());
__Check("7");
