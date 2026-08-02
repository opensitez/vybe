// vybe-test: csharp/csharp_hashcode/class_overriding_get_hash_code_uses_hashcode_combine
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Point{
    public int X,Y;
    public override int GetHashCode()=>System.HashCode.Combine(X,Y);
}
var p1=new Point{X=1,Y=2};
var p2=new Point{X=1,Y=2};
__Check((p1.GetHashCode()==p2.GetHashCode()).ToString(), "True");
