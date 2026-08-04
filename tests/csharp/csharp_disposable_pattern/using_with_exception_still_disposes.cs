// vybe-test: csharp/csharp_disposable_pattern/using_with_exception_still_disposes
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

class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}
var r=new R();
try{using(r){throw new System.Exception();}}catch{}
__P((r.Gone).ToString());
__Check("True");
