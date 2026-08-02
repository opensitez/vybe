// vybe-test: csharp/csharp_disposable_pattern/using_with_exception_still_disposes
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}
var r=new R();
try{using(r){throw new System.Exception();}}catch{}
__Check((r.Gone).ToString(), "True");
