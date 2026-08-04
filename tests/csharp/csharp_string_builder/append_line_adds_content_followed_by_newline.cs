// vybe-test: csharp/csharp_string_builder/append_line_adds_content_followed_by_newline
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

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

var sb = new System.Text.StringBuilder();
sb.AppendLine("line1");
__P((sb.Length > 5).ToString());
__Check("True");
