// vybe-test: csharp/csharp_goto_switch_labels/goto_label_before_switch_merge_point
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int mode = 1;
string result = "";
if (mode == 0) goto merge;
switch (mode) {
    case 1: result += "one"; break;
    case 2: result += "two"; break;
}
merge:
result += "!";
__Check((result).ToString(), "one!");
