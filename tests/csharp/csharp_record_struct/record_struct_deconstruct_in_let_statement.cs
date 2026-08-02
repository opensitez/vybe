// vybe-test: csharp/csharp_record_struct/record_struct_deconstruct_in_let_statement
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Vec(int X,int Y);
var v=new Vec(3,4);
var(x,y)=v;
__Check((x+y).ToString(), "7");
