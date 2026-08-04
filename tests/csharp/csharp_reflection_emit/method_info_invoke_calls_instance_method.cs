// vybe-test: csharp/csharp_reflection_emit/method_info_invoke_calls_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

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

class Adder{public int Add(int a,int b)=>a+b;}
var mi=typeof(Adder).GetMethod("Add");
var result=mi.Invoke(new Adder(),new object[]{3,4});
__P((result).ToString());
__Check("7");
