// vybe-test: csharp/csharp_anonymous_types/anonymous_type_created_with_new_projection_syntax
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new{Name="Alice",Age=30};
__Check((a.Name).ToString(), "Alice"); __Check((a.Age).ToString(), "30");
