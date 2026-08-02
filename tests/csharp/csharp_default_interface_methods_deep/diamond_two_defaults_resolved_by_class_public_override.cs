// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_two_defaults_resolved_by_class_public_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

interface IA{void M()=>Console.WriteLine("A");} interface IB{void M()=>Console.WriteLine("B");} class C:IA,IB{public void M()=>Console.WriteLine("C");} new C().M();
