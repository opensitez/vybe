// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_readonly
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly record struct Size(int W,int H); var s=new Size(2,3); var t=s with{H=8}; __Check((s.H).ToString(), "3"); __Check((t.H).ToString(), "8");
