// vybe-test: csharp/csharp_switch_type_patterns/switch_on_string_matches_exact_literal_case
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Pick(string key) {
    switch (key) {
        case "go": return "G";
        case "stop": return "S";
        default: return "?";
    }
}
__Check((Pick("go")).ToString(), "G");
