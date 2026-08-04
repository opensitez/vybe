// vybe-test: csharp/csharp_array_advanced/array_fill_sets_all_elements_to_value
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

int[] arr=new int[5];
System.Array.Fill(arr,7);
__P((arr[2]).ToString());
__Check("7");
