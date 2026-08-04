// vybe-test: csharp/csharp_goto_labels/goto_jumps_to_labeled_statement
// origin: languages/csharp/tests/csharp/test_csharp_goto_labels.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int i=0;
start:
if(i<5){i++;goto start;}
__P((i).ToString());
__Check("5");
