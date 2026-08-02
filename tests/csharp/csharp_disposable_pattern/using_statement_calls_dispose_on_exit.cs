// vybe-test: csharp/csharp_disposable_pattern/using_statement_calls_dispose_on_exit
// origin: languages/csharp/tests/csharp/test_csharp_disposable_pattern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Resource:System.IDisposable{
    public bool Disposed;
    public void Dispose(){Disposed=true;}
}
var r=new Resource();
using(r){}
__Check((r.Disposed).ToString(), "True");
