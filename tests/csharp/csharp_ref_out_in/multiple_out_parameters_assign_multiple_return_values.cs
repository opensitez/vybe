// vybe-test: csharp/csharp_ref_out_in/multiple_out_parameters_assign_multiple_return_values
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

using static __Harness;

void Split(string s, out string head, out string tail){
    int mid=s.Length/2;
    head=s.Substring(0,mid); tail=s.Substring(mid);
}
Split("abcdef",out string h,out string t);
__P((h).ToString());
__P((t).ToString());
__Check("abc\ndef");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
