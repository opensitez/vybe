// vybe-test: csharp/csharp_value_ref_semantics/class_assignment_creates_alias_not_copy
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pt{public int X,Y;}
var a=new Pt{X=1,Y=2};
var b=a; b.X=99;
__Check((a.X).ToString(), "99");
