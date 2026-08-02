// vybe-test: csharp/csharp_const_and_readonly_fields/const_enum_member_casts_to_underlying_integer_value
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Code { Ok = 0, Err = 1 }
const Code status = Code.Ok;
__Check(((int)status).ToString(), "0");
