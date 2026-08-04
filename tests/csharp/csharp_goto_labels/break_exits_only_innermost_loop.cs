// vybe-test: csharp/csharp_goto_labels/break_exits_only_innermost_loop
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

int count=0;
for(int i=0;i<3;i++){
    for(int j=0;j<3;j++){
        if(j==1) break;
        count++;
    }
}
__P((count).ToString());
__Check("3");
