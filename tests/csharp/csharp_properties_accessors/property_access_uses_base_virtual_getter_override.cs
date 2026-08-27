// vybe-test: csharp/csharp_properties_accessors/property_access_uses_base_virtual_getter_override
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

BasePerson person = new Employee();
__P((person.Label).ToString());
__Check("employee");

class BasePerson {
    public virtual string Label { get { return "base"; } }
}

class Employee : BasePerson {
    public override string Label { get { return "employee"; } }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
