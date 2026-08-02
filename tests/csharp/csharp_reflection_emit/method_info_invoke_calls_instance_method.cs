// vybe-test: csharp/csharp_reflection_emit/method_info_invoke_calls_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Adder{public int Add(int a,int b)=>a+b;}
var mi=typeof(Adder).GetMethod("Add");
var result=mi.Invoke(new Adder(),new object[]{3,4});
__Check((result).ToString(), "7");
