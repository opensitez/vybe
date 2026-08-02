// vybe-test: csharp/csharp_covariance_contravariance/array_covariance_allows_derived_array_in_base_array_reference
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string[] strings = { "a", "b" };
object[] objects = strings;
__Check((objects[0]).ToString(), "a");
