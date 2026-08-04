// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_abstract_class_with_derived
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

namespace Abs;
abstract class Base { public abstract int Get(); }
class Impl : Base { public override int Get() => 4; }
__P((new Impl().Get()).ToString());
__Check("4");
