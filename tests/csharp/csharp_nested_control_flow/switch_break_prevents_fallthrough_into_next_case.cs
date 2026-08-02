// vybe-test: csharp/csharp_nested_control_flow/switch_break_prevents_fallthrough_into_next_case
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int code = 2;
string label = "";
switch (code) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    case 3: label = "three"; break;
}
__Check((label).ToString(), "two");
