// vybe-test: csharp/csharp_reflection_activation/typeof_name_reports_string_type_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

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

__P((typeof(string).Name).ToString());
__Check("String");
