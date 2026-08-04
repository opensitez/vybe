// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_nominal
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

record struct Config{public int Port{get;init;}=80;} var c=new Config{Port=8080}; var d=c with{Port=443}; __P((c.Port).ToString()); __P((d.Port).ToString());
__Check("8080\n443");
