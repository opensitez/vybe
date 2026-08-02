// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_copy_to_array_preserves_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"xy"; byte[] buf=new byte[2]; bytes.CopyTo(buf); __Check((buf[0]).ToString(), "120"); __Check((buf[1]).ToString(), "121");
