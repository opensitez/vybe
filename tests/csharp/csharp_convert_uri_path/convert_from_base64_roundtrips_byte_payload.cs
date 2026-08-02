// vybe-test: csharp/csharp_convert_uri_path/convert_from_base64_roundtrips_byte_payload
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Convert.FromBase64String("AQID");
__Check((bytes.Length).ToString(), "3");
__Check((bytes[2]).ToString(), "3");
