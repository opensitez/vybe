// vybe-test: csharp/csharp_utf8_string_literals/utf8_literal_foreach_sum_of_bytes
// origin: languages/csharp/tests/csharp/test_csharp_utf8_string_literals.rs

var bytes=u8"ab"; int sum=0; foreach(var b in bytes) sum+=b; Console.WriteLine(sum);
