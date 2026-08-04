// vybe-test: csharp/csharp_array_operations/array_clear_fills_range_with_default_values
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

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

int[] a = {1,2,3,4,5};
System.Array.Clear(a, 1, 3);
__P((a[0]).ToString()); __P((a[2]).ToString());
__Check("1\n0");
