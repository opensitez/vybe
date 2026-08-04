// vybe-test: csharp/csharp_array_advanced/array_sort_then_binary_search_finds_element
// origin: languages/csharp/tests/csharp/test_csharp_array_advanced.rs

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

int[] arr={5,3,1,4,2};
System.Array.Sort(arr);
int idx=System.Array.BinarySearch(arr,4);
__P((idx).ToString());
__Check("3");
