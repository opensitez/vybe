// vybe-test: csharp/csharp_ref_return_semantics/ref_return_allows_mutating_array_element_through_alias
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

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

int[] data = { 1, 2, 3 };
ref int Slot(int index) => ref data[index];
ref int cell = ref Slot(1);
cell = 9;
__P((data[1]).ToString());
__Check("9");
