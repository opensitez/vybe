// vybe-test: csharp/csharp_ref_return_semantics/ref_return_allows_mutating_array_element_through_alias
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = { 1, 2, 3 };
ref int Slot(int index) => ref data[index];
ref int cell = ref Slot(1);
cell = 9;
__Check((data[1]).ToString(), "9");
