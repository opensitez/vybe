// vybe-test: csharp/csharp_bit_converter_endian/bit_converter_int32_roundtrip_preserves_numeric_value
// origin: languages/csharp/tests/csharp/test_csharp_bit_converter_endian.rs

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

var bytes = System.BitConverter.GetBytes(1024);
__P((System.BitConverter.ToInt32(bytes, 0)).ToString());
__Check("1024");
