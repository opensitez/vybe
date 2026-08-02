// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_element_type_is_enum
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Kind{A,B} var values=System.Enum.GetValues(typeof(Kind)); __Check((values.GetType().GetElementType().Name).ToString(), "Kind");
