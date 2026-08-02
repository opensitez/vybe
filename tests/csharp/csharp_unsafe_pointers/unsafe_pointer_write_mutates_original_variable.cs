// vybe-test: csharp/csharp_unsafe_pointers/unsafe_pointer_write_mutates_original_variable
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unsafe{
    int x=1;
    int* p=&x;
    *p=99;
    __Check((x).ToString(), "99");
}
