// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_field_mutation
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace State;
class Counter { public int Count; }
var c = new Counter { Count = 1 };
c.Count = 5;
__Check((c.Count).ToString(), "5");
