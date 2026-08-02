// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_enum_via_public_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Status{enum Code{Ok=0,Fail=1} public int Read()=>(int)Code.Ok;} __Check((new Status().Read()).ToString(), "0");
