// vybe-test: csharp/csharp_record_struct/record_struct_with_expression_creates_new_copy
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Point(int X,int Y);
var a=new Point(1,2);
var b=a with{X=99};
__Check((a.X).ToString(), "1"); __Check((b.X).ToString(), "99");
