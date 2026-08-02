// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_abstract_class_with_derived
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Abs;
abstract class Base { public abstract int Get(); }
class Impl : Base { public override int Get() => 4; }
__Check((new Impl().Get()).ToString(), "4");
