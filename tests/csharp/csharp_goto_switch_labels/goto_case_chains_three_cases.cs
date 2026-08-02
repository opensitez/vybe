// vybe-test: csharp/csharp_goto_switch_labels/goto_case_chains_three_cases
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 1;
string r = "";
switch (n) {
    case 1: r += "1"; goto case 2;
    case 2: r += "2"; goto case 3;
    case 3: r += "3"; break;
}
__Check((r).ToString(), "123");
