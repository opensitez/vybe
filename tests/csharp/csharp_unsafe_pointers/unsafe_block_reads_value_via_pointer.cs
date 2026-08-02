// vybe-test: csharp/csharp_unsafe_pointers/unsafe_block_reads_value_via_pointer
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unsafe{
    int x=42;
    int* p=&x;
    __Check((*p).ToString(), "42");
}
