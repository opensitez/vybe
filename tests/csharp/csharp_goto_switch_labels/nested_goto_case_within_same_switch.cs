// vybe-test: csharp/csharp_goto_switch_labels/nested_goto_case_within_same_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 1;
string s = "";
switch (x) {
    case 1: s += "a"; goto case 2;
    case 2: s += "b"; goto case 3;
    case 3: s += "c"; break;
}
__Check((s).ToString(), "abc");
