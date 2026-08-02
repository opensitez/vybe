// vybe-test: csharp/csharp_generics_advanced/generic_list_works_with_interface_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

interface IAnimal { string Sound(); }
class Cat : IAnimal { public string Sound() => "meow"; }
var animals = new System.Collections.Generic.List<IAnimal> { new Cat() };
foreach(var a in animals) Console.WriteLine(a.Sound());
