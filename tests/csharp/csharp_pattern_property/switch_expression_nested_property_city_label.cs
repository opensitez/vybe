// vybe-test: csharp/csharp_pattern_property/switch_expression_nested_property_city_label
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Address { public string City; } class Person { public Address Addr; } string Where(object p)=>p switch{Person{Addr:{City:"NYC"}}=>"metro",_=>"other"}; __P((Where(new Person{Addr=new Address{City="NYC"}})).ToString());
__Check("metro");
