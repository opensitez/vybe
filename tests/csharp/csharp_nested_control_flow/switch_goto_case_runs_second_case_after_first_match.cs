// vybe-test: csharp/csharp_nested_control_flow/switch_goto_case_runs_second_case_after_first_match
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

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
