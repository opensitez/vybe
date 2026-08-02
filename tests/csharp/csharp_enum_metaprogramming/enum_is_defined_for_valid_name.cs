// vybe-test: csharp/csharp_enum_metaprogramming/enum_is_defined_for_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Phase{Start,End} __Check((System.Enum.IsDefined(typeof(Phase),"Start")).ToString(), "True");
