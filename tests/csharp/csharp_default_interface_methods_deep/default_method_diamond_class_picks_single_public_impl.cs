// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_diamond_class_picks_single_public_impl
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

interface IA{void Print()=>Console.WriteLine("A");} interface IB{void Print()=>Console.WriteLine("B");} class Merge:IA,IB{public void Print()=>Console.WriteLine("M");} new Merge().Print();
