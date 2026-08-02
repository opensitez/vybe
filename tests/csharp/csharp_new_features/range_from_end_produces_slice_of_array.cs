// vybe-test: csharp/csharp_new_features/range_from_end_produces_slice_of_array
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = {1,2,3,4,5};
var last2 = arr[^2..];
__Check((last2[0]).ToString(), "4"); __Check((last2[1]).ToString(), "5");
