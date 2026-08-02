// vybe-test: csharp/csharp_goto_switch_labels/goto_case_preserves_order_with_break
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int k = 1;
string buf = "";
switch (k) {
    case 1: buf += "1"; goto case 2;
    case 2: buf += "2"; break;
    case 3: buf += "3"; break;
}
__Check((buf).ToString(), "12");
