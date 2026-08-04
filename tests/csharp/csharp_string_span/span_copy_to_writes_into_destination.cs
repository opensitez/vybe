// vybe-test: csharp/csharp_string_span/span_copy_to_writes_into_destination
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

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

int[] src={1,2,3};
int[] dst=new int[3];
src.AsSpan().CopyTo(dst);
__P((dst[2]).ToString());
__Check("3");
