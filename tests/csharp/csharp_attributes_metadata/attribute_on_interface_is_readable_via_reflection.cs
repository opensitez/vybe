// vybe-test: csharp/csharp_attributes_metadata/attribute_on_interface_is_readable_via_reflection
// origin: languages/csharp/tests/csharp/test_csharp_attributes_metadata.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [AttributeUsage(AttributeTargets.Interface)] class ContractAttribute : Attribute { public string Name { get; } public ContractAttribute(string name) { Name = name; } } [Contract("service")] interface IService { } var attr = (ContractAttribute)Attribute.GetCustomAttribute(typeof(IService), typeof(ContractAttribute)); __Check((attr.Name).ToString(), "service");
