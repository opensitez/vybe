// vybe-test: csharp/csharp_ref_out_in/ref_local_aliases_array_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={1,2,3};
ref int second=ref arr[1];
second=99;
__Check((arr[1]).ToString(), "99");
