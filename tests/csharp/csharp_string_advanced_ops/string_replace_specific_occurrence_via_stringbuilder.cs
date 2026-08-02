// vybe-test: csharp/csharp_string_advanced_ops/string_replace_specific_occurrence_via_stringbuilder
// origin: languages/csharp/tests/csharp/test_csharp_string_advanced_ops.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="aababc";
var sb=new System.Text.StringBuilder(s);
int idx=s.IndexOf("ab",1);
sb.Remove(idx,2).Insert(idx,"XX");
__Check((sb.ToString()).ToString(), "aaXXbc");
