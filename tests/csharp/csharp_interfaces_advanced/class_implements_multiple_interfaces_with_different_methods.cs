// vybe-test: csharp/csharp_interfaces_advanced/class_implements_multiple_interfaces_with_different_methods
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

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

interface IRead{string Read();}
interface IWrite{void Write(string v);}
class Buffer:IRead,IWrite{
    string _val="";
    public string Read()=>_val;
    public void Write(string v){_val=v;}
}
var b=new Buffer();
((IWrite)b).Write("data");
__P((((IRead)b).Read()).ToString());
__Check("data");
