// vybe-test: csharp/csharp_loops/break_exits_innermost_loop_only
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

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

int total = 0;
for(int i=0;i<3;i++) {
    for(int j=0;j<3;j++) {
        if(j==1) break;
        total++;
    }
}
__P((total).ToString());
__Check("3");
