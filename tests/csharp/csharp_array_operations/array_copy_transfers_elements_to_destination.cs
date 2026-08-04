// vybe-test: csharp/csharp_array_operations/array_copy_transfers_elements_to_destination
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

int[] src = {10,20,30}; int[] dst = new int[3];
System.Array.Copy(src, dst, 3);
__P((dst[1]).ToString());
__Check("20");
