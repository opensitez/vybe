// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_interface_constraint_on_type_param
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IHasLabel<T> where T:IHasLabel<T>{static abstract string Label();}
struct Tag:IHasLabel<Tag>{public static string Label()=>"tag";}
string Read<T>() where T:IHasLabel<T>=>T.Label(); __Check((Read<Tag>()).ToString(), "tag");
