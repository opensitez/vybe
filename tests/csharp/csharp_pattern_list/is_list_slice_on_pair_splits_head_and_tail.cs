// vybe-test: csharp/csharp_pattern_list/is_list_slice_on_pair_splits_head_and_tail
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data = new[]{5,6}; if (data is [var a,..var rest]) __Check((a+rest[0]).ToString(), "11");
