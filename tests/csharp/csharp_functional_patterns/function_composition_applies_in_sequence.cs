// vybe-test: csharp/csharp_functional_patterns/function_composition_applies_in_sequence
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

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

System.Func<int,int> triple=x=>x*3;
System.Func<int,int> addOne=x=>x+1;
var composed=new[]{1,2,3}.Select(triple).Select(addOne);
foreach(var n in composed) __P((n).ToString());
__Check("4\n7\n10");
