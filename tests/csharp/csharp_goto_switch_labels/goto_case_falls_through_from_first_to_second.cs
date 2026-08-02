// vybe-test: csharp/csharp_goto_switch_labels/goto_case_falls_through_from_first_to_second
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int code = 1;
string trace = "";
switch (code) {
    case 1:
        trace += "A";
        goto case 2;
    case 2:
        trace += "B";
        break;
}
__Check((trace).ToString(), "AB");
