// vybe-test: csharp/csharp_pattern_matching/tuple_pattern_deconstructs_two_element_tuple
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var point = (1, 0);
string axis = point switch {
    (0, 0) => "origin",
    (_, 0) => "x-axis",
    (0, _) => "y-axis",
    _       => "other"
};
__Check((axis).ToString(), "x-axis");
