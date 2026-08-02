// vybe-test: csharp/csharp_anonymous_types/anonymous_type_property_names_inferred_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int id=7; string name="Bob";
var obj=new{id,name};
__Check((obj.id).ToString(), "7"); __Check((obj.name).ToString(), "Bob");
