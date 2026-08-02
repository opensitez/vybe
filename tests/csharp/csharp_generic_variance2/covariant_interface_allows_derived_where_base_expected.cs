// vybe-test: csharp/csharp_generic_variance2/covariant_interface_allows_derived_where_base_expected
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IReader<out T>{T Read();}
class StringReader:IReader<string>{public string Read()=>"hello";}
IReader<object> r=new StringReader();
__Check((r.Read()).ToString(), "hello");
