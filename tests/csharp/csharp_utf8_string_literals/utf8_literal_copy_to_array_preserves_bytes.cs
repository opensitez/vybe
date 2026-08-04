// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_copy_to_array_preserves_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var bytes=u8"xy"; byte[] buf=new byte[2]; bytes.CopyTo(buf); __P((buf[0]).ToString()); __P((buf[1]).ToString());
__Check("120\n121");
