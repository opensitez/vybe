// vybe-test: csharp/csharp_ranges_indices/index_variable_used_in_array_access
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

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

int[] a={10,20,30,40,50};
System.Index i=^2;
__P((a[i]).ToString());
__Check("40");
