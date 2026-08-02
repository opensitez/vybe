// vybe-test: csharp/csharp_covariance_contravariance/contravariant_in_allows_base_generic_argument_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IWriter<in T> { void Write(T value); }
class ObjectWriter : IWriter<object> {
    public void Write(object value) => __Check((value).ToString(), "typed");
}
IWriter<string> writer = new ObjectWriter();
writer.Write("typed");
