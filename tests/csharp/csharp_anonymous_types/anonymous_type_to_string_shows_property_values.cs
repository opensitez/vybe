// vybe-test: csharp/csharp_anonymous_types/anonymous_type_to_string_shows_property_values
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new{X=3,Y=4};
__Check((a.ToString().Contains("X = 3")).ToString(), "True");
