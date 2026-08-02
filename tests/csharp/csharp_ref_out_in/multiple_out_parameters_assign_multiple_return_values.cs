// vybe-test: csharp/csharp_ref_out_in/multiple_out_parameters_assign_multiple_return_values
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

void Split(string s, out string head, out string tail){
    int mid=s.Length/2;
    head=s.Substring(0,mid); tail=s.Substring(mid);
}
Split("abcdef",out string h,out string t);
__Check((h).ToString(), "abc"); __Check((t).ToString(), "def");
