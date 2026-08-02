// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_with_verbatim_literal_segment_beside_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var drive = "C"; __Check(($@"{drive}\temp").ToString(), "C\\temp");
