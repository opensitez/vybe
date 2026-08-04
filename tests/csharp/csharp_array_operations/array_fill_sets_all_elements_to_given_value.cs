// vybe-test: csharp/csharp_array_operations/array_fill_sets_all_elements_to_given_value
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

int[] a = new int[4];
System.Array.Fill(a, 7);
__P((a[3]).ToString());
__Check("7");
