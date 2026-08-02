// vybe-test: csharp/csharp_goto_labels/goto_jumps_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int i=0;
start:
if(i<5){i++;goto start;}
__Check((i).ToString(), "5");
