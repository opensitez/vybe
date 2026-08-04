// vybe-test: csharp/csharp_interface_contracts/idisposable_dispose_called_by_using_statement
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

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

class Resource : System.IDisposable {
    public static int Disposed = 0;
    public void Dispose() => Disposed++;
}
using(var r = new Resource()) { }
__P((Resource.Disposed).ToString());
__Check("1");
