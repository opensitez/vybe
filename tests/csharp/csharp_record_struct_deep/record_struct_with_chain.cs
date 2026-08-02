// vybe-test: csharp/csharp_record_struct_deep/record_struct_with_chain
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Box(int W,int H); var a=new Box(1,1); var b=a with{W=2}; var c=b with{H=3}; __Check((a.W).ToString(), "1"); __Check((c.W).ToString(), "2"); __Check((c.H).ToString(), "3");
