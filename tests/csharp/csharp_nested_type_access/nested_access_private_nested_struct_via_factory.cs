// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_struct_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Builder{struct Part{public int N;} Part Make(){return new Part{N=8};} public int Build()=>Make().N;} __Check((new Builder().Build()).ToString(), "8");
