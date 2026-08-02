// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_char_ptr_reads_first_character_code
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char[] arr={'A','B'}; unsafe{fixed(char* ptr=&arr[0]){__Check((*ptr).ToString(), "65");}}
