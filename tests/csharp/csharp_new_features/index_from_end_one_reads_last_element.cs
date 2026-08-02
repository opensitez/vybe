// vybe-test: csharp/csharp_new_features/index_from_end_one_reads_last_element
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = {10,20,30};
__Check((arr[^1]).ToString(), "30");
