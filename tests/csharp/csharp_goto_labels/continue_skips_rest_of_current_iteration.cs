// vybe-test: csharp/csharp_goto_labels/continue_skips_rest_of_current_iteration
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

int sum=0;
for(int i=1;i<=10;i++){
    if(i%2==0) continue;
    sum+=i;
}
__P((sum).ToString());
__Check("25");
