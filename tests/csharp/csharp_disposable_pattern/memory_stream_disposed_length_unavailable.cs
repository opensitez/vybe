// vybe-test: csharp/csharp_disposable_pattern/memory_stream_disposed_length_unavailable
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.IO.MemoryStream ms;
using(ms=new System.IO.MemoryStream()){}
string r="";
try{var _=ms.Length;}catch(System.ObjectDisposedException){r="disposed";}
__Check((r).ToString(), "disposed");
