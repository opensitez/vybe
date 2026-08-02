// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_visibility_public_from_external
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Api{public class Endpoint{public string Path="/v1";}} __Check((new Api.Endpoint().Path).ToString(), "/v1");
