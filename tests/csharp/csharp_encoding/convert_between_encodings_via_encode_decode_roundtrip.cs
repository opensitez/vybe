// vybe-test: csharp/csharp_encoding/convert_between_encodings_via_encode_decode_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = "test";
byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
string result = System.Text.Encoding.UTF8.GetString(bytes);
__Check((text == result).ToString(), "True");
