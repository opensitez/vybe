// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_to_string
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mode{public enum Kind{Alpha,Beta} public string Label()=>Kind.Beta.ToString();} __Check((new Mode().Label()).ToString(), "Beta");
