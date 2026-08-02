// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_operator_overload
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Ops;
struct V { public int N; public static V operator +(V a, V b) => new V { N = a.N + b.N }; }
var r = new V { N = 2 } + new V { N = 3 };
__Check((r.N).ToString(), "5");
