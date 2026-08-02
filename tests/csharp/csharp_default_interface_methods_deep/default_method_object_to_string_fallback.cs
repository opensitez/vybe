// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_object_to_string_fallback
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IObj{string Desc()=>ToString();} class Thing:IObj{public override string ToString()=>"thing";} __Check((new Thing().Desc()).ToString(), "thing");
