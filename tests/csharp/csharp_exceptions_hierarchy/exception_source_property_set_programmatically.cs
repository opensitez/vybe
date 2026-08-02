// vybe-test: csharp/csharp_exceptions_hierarchy/exception_source_property_set_programmatically
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var ex=new System.Exception("e");
ex.Source="MyModule";
__Check((ex.Source).ToString(), "MyModule");
