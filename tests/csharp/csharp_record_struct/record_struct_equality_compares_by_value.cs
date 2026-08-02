// vybe-test: csharp/csharp_record_struct/record_struct_equality_compares_by_value
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Color(int R,int G,int B);
var c1=new Color(255,0,0); var c2=new Color(255,0,0);
__Check((c1==c2).ToString(), "True");
