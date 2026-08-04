// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_operator_overload
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

namespace Ops;
struct V { public int N; public static V operator +(V a, V b) => new V { N = a.N + b.N }; }
var r = new V { N = 2 } + new V { N = 3 };
__P((r.N).ToString());
__Check("5");
