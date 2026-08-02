// vybe-test: csharp/type_gettype_dynamic/get_type_returns_null_for_an_unknown_name
// origin: languages/csharp/tests/csharp/test_type_gettype_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Program {
    static void Main() {
        System.__Check((System.Type.GetType("NoSuchTypeAnywhere") == null).ToString(), "True");
    }
}
