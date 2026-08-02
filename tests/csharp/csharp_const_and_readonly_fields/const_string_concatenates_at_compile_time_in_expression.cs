// vybe-test: csharp/csharp_const_and_readonly_fields/const_string_concatenates_at_compile_time_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Labels {
    public const string Base = "user";
    public const string Full = Base + "_id";
}
__Check((Labels.Full).ToString(), "user_id");
