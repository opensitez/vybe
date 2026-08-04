// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_interface_implementation
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

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

namespace App;
interface IRun { string Go(); }
class Runner : IRun { public string Go() => "go"; }
IRun r = new Runner();
__P((r.Go()).ToString());
__Check("go");
