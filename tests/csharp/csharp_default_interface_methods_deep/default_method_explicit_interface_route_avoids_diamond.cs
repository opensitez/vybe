// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_explicit_interface_route_avoids_diamond
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

interface IA{void Show()=>Console.WriteLine("A");} interface IB{void Show()=>Console.WriteLine("B");} class Split:IA,IB{void IA.Show()=>Console.WriteLine("IA"); void IB.Show()=>Console.WriteLine("IB");} ((IA)new Split()).Show();
