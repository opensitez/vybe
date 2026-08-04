// vybe-test: csharp/csharp_text_stringbuilder/string_builder_capacity_grows_automatically
// origin: languages/csharp/tests/csharp/test_csharp_text_stringbuilder.rs

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

var sb=new System.Text.StringBuilder(4);
for(int i=0;i<100;i++) sb.Append('x');
__P((sb.Length).ToString());
__Check("100");
