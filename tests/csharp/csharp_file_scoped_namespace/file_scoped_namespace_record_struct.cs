// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_record_struct
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Geo;
record struct Point(int X, int Y);
__Check((new Point(3, 4).Y).ToString(), "4");
