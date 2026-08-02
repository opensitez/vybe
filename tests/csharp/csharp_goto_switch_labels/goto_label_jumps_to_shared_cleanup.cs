// vybe-test: csharp/csharp_goto_switch_labels/goto_label_jumps_to_shared_cleanup
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 1;
string msg = "";
if (n == 1) goto cleanup;
msg = "skip";
cleanup:
msg = "ok";
__Check((msg).ToString(), "ok");
