// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_concatenated_with_plain_string
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var id = 7; __Check(("id=" + $"{id}").ToString(), "id=7");
