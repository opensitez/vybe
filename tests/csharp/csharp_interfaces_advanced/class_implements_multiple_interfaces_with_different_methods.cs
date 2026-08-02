// vybe-test: csharp/csharp_interfaces_advanced/class_implements_multiple_interfaces_with_different_methods
// origin: languages/csharp/tests/csharp/test_csharp_interfaces_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((((IRead)b).Read()).ToString(), "data");
