// vybe-test: csharp/csharp_goto_switch_labels/goto_label_skips_else_branch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int pick = 1;
string r = "";
if (pick == 0) r = "zero";
else goto show;
show:
r = "one";
__Check((r).ToString(), "one");
