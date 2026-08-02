// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_two_defaults_resolved_by_explicit_interface_impl
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

interface IA{void M()=>Console.WriteLine("A");} interface IB{void M()=>Console.WriteLine("B");} class C:IA,IB{void IA.M()=>Console.WriteLine("IA"); void IB.M()=>Console.WriteLine("IB");} ((IA)new C()).M(); ((IB)new C()).M();
