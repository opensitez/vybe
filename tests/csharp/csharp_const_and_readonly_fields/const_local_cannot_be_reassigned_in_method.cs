// vybe-test: csharp/csharp_const_and_readonly_fields/const_local_cannot_be_reassigned_in_method
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

const int step = 5;
__Check((step * 2).ToString(), "10");
