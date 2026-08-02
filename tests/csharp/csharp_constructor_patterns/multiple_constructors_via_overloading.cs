// vybe-test: csharp/csharp_constructor_patterns/multiple_constructors_via_overloading
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Range{public int Lo,Hi;
    public Range():this(0,100){}
    public Range(int lo,int hi){Lo=lo;Hi=hi;}
}
var r1=new Range(); var r2=new Range(5,10);
__Check((r1.Lo).ToString(), "0"); __Check((r2.Hi).ToString(), "10");
