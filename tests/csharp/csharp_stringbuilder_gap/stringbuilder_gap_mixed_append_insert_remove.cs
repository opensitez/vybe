// vybe-test: csharp/csharp_stringbuilder_gap/stringbuilder_gap_mixed_append_insert_remove
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

var sb=new System.Text.StringBuilder("start"); sb.Append("_end"); sb.Insert(5,"-mid-"); sb.Remove(0,6); __P((sb.ToString()).ToString());
__Check("mid-_end");
