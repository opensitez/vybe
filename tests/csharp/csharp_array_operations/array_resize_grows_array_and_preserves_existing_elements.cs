// vybe-test: csharp/csharp_array_operations/array_resize_grows_array_and_preserves_existing_elements
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

int[] a = {1,2,3};
System.Array.Resize(ref a, 5);
__P((a.Length).ToString()); __P((a[2]).ToString());
__Check("5\n3");
