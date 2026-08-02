// vybe-test: csharp/csharp_object_initializers/anonymous_type_initializer_infers_property_names
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name="Alice"; int age=30;
var anon=new{name,age};
__Check((anon.name).ToString(), "Alice"); __Check((anon.age).ToString(), "30");
