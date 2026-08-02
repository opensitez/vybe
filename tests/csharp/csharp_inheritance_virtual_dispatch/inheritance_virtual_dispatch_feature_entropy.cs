// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch:71"; __Check((feature.Length >= 1).ToString(), "True");
