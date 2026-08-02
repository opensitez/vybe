// vybe-test: csharp/csharp_switch_type_patterns/switch_on_int_falls_through_to_default_when_no_case_matches
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int code = 99;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    default: label = "other"; break;
}
__Check((label).ToString(), "other");
