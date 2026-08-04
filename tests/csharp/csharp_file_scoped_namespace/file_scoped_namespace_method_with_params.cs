// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_method_with_params
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

namespace Calc;
class Adder { public int Sum(int a, int b) => a + b; }
__P((new Adder().Sum(2, 3)).ToString());
__Check("5");
