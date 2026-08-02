// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_class_deconstruct_two
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Size{public int W,H; public void Deconstruct(out int w,out int h){w=W;h=H;}} var (w,h)=new Size{W=3,H=4}; __Check((w+h).ToString(), "7");
