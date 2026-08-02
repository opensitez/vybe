// vybe-test: csharp/csharp_covariance_contravariance/covariant_out_allows_derived_generic_argument_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IReader<out T> { T Read(); }
class StringReader : IReader<string> {
    public string Read() => "hello";
}
IReader<object> reader = new StringReader();
__Check((reader.Read()).ToString(), "hello");
