// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_chain
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

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

record struct Box(int W,int H); var a=new Box(1,1); var b=a with{W=2}; var c=b with{H=3}; __P((a.W).ToString()); __P((c.W).ToString()); __P((c.H).ToString());
__Check("1\n2\n3");
