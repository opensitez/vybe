// vybe-test: csharp/csharp_string_builder/append_format_interpolates_value_into_buffer
// origin: languages/csharp/tests/csharp/test_csharp_string_builder.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var sb = new System.Text.StringBuilder();
sb.AppendFormat("Value={0}", 42);
__Check((sb.ToString()).ToString(), "Value=42");
