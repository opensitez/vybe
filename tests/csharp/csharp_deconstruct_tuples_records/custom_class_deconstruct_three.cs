// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_class_deconstruct_three
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box{public int A,B,C; public void Deconstruct(out int a,out int b,out int c){a=A;b=B;c=C;}} var (a,b,c)=new Box{A=1,B=2,C=3}; __Check((a+b+c).ToString(), "6");
