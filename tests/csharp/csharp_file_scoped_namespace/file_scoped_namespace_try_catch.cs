// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_try_catch
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Err;
string Read() { try { return "ok"; } catch { return "bad"; } }
__Check((Read()).ToString(), "ok");
