// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch"; __Check((feature[0] == feature[0]).ToString(), "True");
