// vybe-test: csharp/csharp_goto_switch_labels/goto_default_from_non_matching_case_value
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 99;
string label = "";
switch (n) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    default:
        label = "other";
        break;
}
__Check((label).ToString(), "other");
