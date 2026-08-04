// vybe-test: csharp/csharp_string_advanced_ops/string_replace_specific_occurrence_via_stringbuilder
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

string s="aababc";
var sb=new System.Text.StringBuilder(s);
int idx=s.IndexOf("ab",1);
sb.Remove(idx,2).Insert(idx,"XX");
__P((sb.ToString()).ToString());
__Check("aaXXbc");
