// vybe-test: csharp/csharp_caller_info_attributes/caller_line_number_from_loop_body
// origin: languages/csharp/tests/csharp/test_csharp_caller_info_attributes.rs

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

class Trace {
    public static void Show([System.Runtime.CompilerServices.CallerLineNumber] int line = 0) => __P((line).ToString());
}
for (int i = 0; i < 1; i++) Trace.Show();
__Check("4");
