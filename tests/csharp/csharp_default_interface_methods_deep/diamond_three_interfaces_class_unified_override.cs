// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_three_interfaces_class_unified_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

interface IA{void P()=>Console.WriteLine("A");} interface IB{void P()=>Console.WriteLine("B");} interface IC{void P()=>Console.WriteLine("C");} class U:IA,IB,IC{public void P()=>Console.WriteLine("U");} new U().P();
