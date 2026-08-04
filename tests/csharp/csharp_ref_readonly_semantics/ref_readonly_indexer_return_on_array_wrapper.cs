// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_indexer_return_on_array_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

class Buffer{private int[] _data={5,6,7}; public ref readonly int this[int i]=>ref _data[i];} var b=new Buffer(); __P((b[2]).ToString());
__Check("7");
