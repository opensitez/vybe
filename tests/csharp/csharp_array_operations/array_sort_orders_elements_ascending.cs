// vybe-test: csharp/csharp_array_operations/array_sort_orders_elements_ascending
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

int[] a = {3,1,4,1,5};
System.Array.Sort(a);
__P((a[0]).ToString()); __P((a[4]).ToString());
__Check("1\n5");
