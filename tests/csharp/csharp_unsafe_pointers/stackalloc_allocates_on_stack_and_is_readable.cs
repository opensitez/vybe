// vybe-test: csharp/csharp_unsafe_pointers/stackalloc_allocates_on_stack_and_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unsafe{
    int* buf=stackalloc int[3]{1,2,3};
    __Check((buf[2]).ToString(), "3");
}
