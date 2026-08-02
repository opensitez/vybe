// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_three_string_fields
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Street { public string Name; } class Addr { public Street S; } class Person { public Addr A; } object p=new Person{A=new Addr{S=new Street{Name="Main"}}}; __Check((p is Person{A:{S:{Name:"Main"}}}).ToString(), "True");
