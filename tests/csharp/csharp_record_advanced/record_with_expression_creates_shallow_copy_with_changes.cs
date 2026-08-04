// vybe-test: csharp/csharp_record_advanced/record_with_expression_creates_shallow_copy_with_changes
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

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

record Config(int Port,string Host);
var c1=new Config(80,"localhost");
var c2=c1 with{Port=443};
__P((c1.Port).ToString()); __P((c2.Port).ToString());
__P((c2.Host).ToString());
__Check("80\n443\nlocalhost");
