// vybe-test: csharp/csharp_string_span/span_of_int_slice_modifies_original_array
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

int[] arr={1,2,3,4,5};
System.Span<int> s=arr.AsSpan(1,3);
s[0]=99;
__P((arr[1]).ToString());
__Check("99");
