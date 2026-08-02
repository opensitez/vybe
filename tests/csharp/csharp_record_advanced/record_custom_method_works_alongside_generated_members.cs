// vybe-test: csharp/csharp_record_advanced/record_custom_method_works_alongside_generated_members
// origin: languages/csharp/tests/csharp/test_csharp_record_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Circle(double Radius){
    public double Area=>System.Math.PI*Radius*Radius;
}
var c=new Circle(1.0);
__Check((c.Area>3.1&&c.Area<3.2).ToString(), "True");
