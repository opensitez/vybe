// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_lambda_in_method
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

namespace Lambda;
class Fn { public int Run() { System.Func<int, int> f = x => x + 1; return f(3); } }
__P((new Fn().Run()).ToString());
__Check("4");
