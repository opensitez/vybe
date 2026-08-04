// vybe-test: csharp/csharp_ref_out_in/multiple_out_parameters_assign_multiple_return_values
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

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

void Split(string s, out string head, out string tail){
    int mid=s.Length/2;
    head=s.Substring(0,mid); tail=s.Substring(mid);
}
Split("abcdef",out string h,out string t);
__P((h).ToString()); __P((t).ToString());
__Check("abc\ndef");
