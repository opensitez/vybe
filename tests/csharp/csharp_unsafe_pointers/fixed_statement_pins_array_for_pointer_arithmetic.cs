// vybe-test: csharp/csharp_unsafe_pointers/fixed_statement_pins_array_for_pointer_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_unsafe_pointers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] arr={10,20,30};
unsafe{
    fixed(int* p=arr){
        __Check((*(p+1)).ToString(), "20");
    }
}
