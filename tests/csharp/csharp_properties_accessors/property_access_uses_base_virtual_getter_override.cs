// vybe-test: csharp/csharp_properties_accessors/property_access_uses_base_virtual_getter_override
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class BasePerson {
    public virtual string Label { get { return "base"; } }
}
class Employee : BasePerson {
    public override string Label { get { return "employee"; } }
}
BasePerson person = new Employee();
__Check((person.Label).ToString(), "employee");
