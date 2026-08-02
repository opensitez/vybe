// vybe-test: csharp/csharp_linq_complex/select_many_flattens_nested_lists
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{
    new[]{1,2,3},
    new[]{4,5},
    new[]{6}
};
int sum=data.SelectMany(x=>x).Sum();
__Check((sum).ToString(), "21");
