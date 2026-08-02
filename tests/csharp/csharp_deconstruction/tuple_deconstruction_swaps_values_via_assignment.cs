// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_swaps_values_via_assignment
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int left = 1;
int right = 2;
(left, right) = (right, left);
__Check((left).ToString(), "2");
__Check((right).ToString(), "1");
