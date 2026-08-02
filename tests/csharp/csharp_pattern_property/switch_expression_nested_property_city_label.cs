// vybe-test: csharp/csharp_pattern_property/switch_expression_nested_property_city_label
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City; } class Person { public Address Addr; } string Where(object p)=>p switch{Person{Addr:{City:"NYC"}}=>"metro",_=>"other"}; __Check((Where(new Person{Addr=new Address{City="NYC"}})).ToString(), "metro");
