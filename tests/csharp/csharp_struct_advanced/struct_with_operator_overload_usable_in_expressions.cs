// vybe-test: csharp/csharp_struct_advanced/struct_with_operator_overload_usable_in_expressions
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Fraction{
    public int Num,Den;
    public static Fraction operator+(Fraction a,Fraction b)=>
        new Fraction{Num=a.Num*b.Den+b.Num*a.Den,Den=a.Den*b.Den};
    public override string ToString()=>$"{Num}/{Den}";
}
var r=new Fraction{Num=1,Den=2}+new Fraction{Num=1,Den=3};
__Check((r.Num).ToString(), "5"); __Check((r.Den).ToString(), "6");
