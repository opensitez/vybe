// vybe-test: csharp/csharp_generic_variance2/contravariant_interface_allows_base_where_derived_expected
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IWriter<in T>{void Write(T v);}
class ObjectWriter:IWriter<object>{public void Write(object v)=>__Check((v).ToString(), "hi");}
IWriter<string> w=new ObjectWriter();
w.Write("hi");
