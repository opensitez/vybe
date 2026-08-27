// vybe-test: csharp/csharp_disposable_pattern/memory_stream_disposed_length_unavailable
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

using static __Harness;

System.IO.MemoryStream ms;
using(ms=new System.IO.MemoryStream()){}
string r="";
try{var _=ms.Length;}
catch(System.ObjectDisposedException){r="disposed";}
__P((r).ToString());
__Check("disposed");

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
