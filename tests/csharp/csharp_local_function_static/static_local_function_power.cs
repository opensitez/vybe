// vybe-test: csharp/csharp_local_function_static/static_local_function_power
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

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

int Pow(int b,int e){static int Loop(int base,int exp,int acc)=>exp==0?acc:Loop(base,exp-1,acc*base); return Loop(b,e,1);} __P((Pow(2,4)).ToString());
__Check("16");
