// vybe-test: csharp/csharp_switch_type_patterns/switch_on_bool_has_separate_true_and_false_arms
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool ok = false;
string label = ok switch { true => "yes", false => "no" };
__Check((label).ToString(), "no");
