// vybe-test: csharp/csharp_pattern_switch_advanced/list_pattern_matches_exact_element_count
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Check(int[] a)=>a switch{
    []=>"empty",
    [_]=>"single",
    [_,_]=>"pair",
    _=>"many"};
__Check((Check(new int[]{})).ToString(), "empty");
__Check((Check(new[]{1})).ToString(), "single");
__Check((Check(new[]{1,2})).ToString(), "pair");
__Check((Check(new[]{1,2,3})).ToString(), "many");
