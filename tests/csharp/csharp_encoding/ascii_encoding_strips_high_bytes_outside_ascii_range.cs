// vybe-test: csharp/csharp_encoding/ascii_encoding_strips_high_bytes_outside_ascii_range
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes = System.Text.Encoding.ASCII.GetBytes("ABC");
__Check((bytes[0]).ToString(), "65");
