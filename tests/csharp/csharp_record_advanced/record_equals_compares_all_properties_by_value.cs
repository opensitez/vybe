// vybe-test: csharp/csharp_record_advanced/record_equals_compares_all_properties_by_value
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

record Pair(int A,int B);
var p1=new Pair(1,2); var p2=new Pair(1,2); var p3=new Pair(1,3);
__P((p1==p2).ToString());
__P((p1==p3).ToString());
__Check("True\nFalse");
