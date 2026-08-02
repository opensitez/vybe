// vybe-test: csharp/csharp_new_features/conditional_ref_var_skips_copy_of_large_struct
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr = {1,2,3};
ref int val = ref arr[1];
val = 99;
__Check((arr[1]).ToString(), "99");
