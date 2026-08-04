// vybe-test: csharp/csharp_disposable_pattern/memory_stream_disposed_length_unavailable
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

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

System.IO.MemoryStream ms;
using(ms=new System.IO.MemoryStream()){}
string r="";
try{var _=ms.Length;}catch(System.ObjectDisposedException){r="disposed";}
__P((r).ToString());
__Check("disposed");
