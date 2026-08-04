// vybe-test: csharp/csharp_random_random/random_next_double_in_zero_one
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

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

var rng=new System.Random(7);
for(int i=0;i<100;i++){
    double v=rng.NextDouble();
    if(v<0.0||v>=1.0){__P(("fail").ToString());return;}
}
__P(("pass").ToString());
__Check("pass");
