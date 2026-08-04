// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_appendline_chained_three
// origin: languages/csharp/tests/csharp/test_csharp_stringbuilder_gap.rs

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

var sb=new System.Text.StringBuilder(); sb.AppendLine("a").AppendLine("b").AppendLine("c"); __P((sb.ToString().Replace("\r\n","\n").Trim().Split('\n').Length).ToString());
__Check("3");
