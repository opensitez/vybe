// vybe-test: csharp/csharp_enum_metaprogramming/enum_cast_to_long_underlying
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Wide:long{Big=10000000000L} __Check(((long)Wide.Big).ToString(), "10000000000");
