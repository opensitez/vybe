// vybe-test: csharp/csharp_reflection_activation/interface_list_contains_declared_contract_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using System.Linq; interface IRun { } class Worker : IRun { } var names = typeof(Worker).GetInterfaces().Select(i => i.Name); foreach (var name in names) Console.WriteLine(name);
