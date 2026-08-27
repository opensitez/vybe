// vybe-test: csharp/csharp_attributes_metadata/attribute_on_interface_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

using static __Harness;
using System;

var attr = (ContractAttribute)Attribute.GetCustomAttribute(typeof(IService), typeof(ContractAttribute));
__P((attr.Name).ToString());
__Check("service");

[AttributeUsage(AttributeTargets.Interface)] class ContractAttribute : Attribute { public string Name { get; } public ContractAttribute(string name) { Name = name; } }

[Contract("service")] interface IService { }

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
