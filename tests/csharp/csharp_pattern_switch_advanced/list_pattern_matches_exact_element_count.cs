// vybe-test: csharp/csharp_pattern_switch_advanced/list_pattern_matches_exact_element_count
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

string Check(int[] a)=>a switch{
    []=>"empty",
    [_]=>"single",
    [_,_]=>"pair",
    _=>"many"};
__P((Check(new int[]{})).ToString());
__P((Check(new[]{1})).ToString());
__P((Check(new[]{1,2})).ToString());
__P((Check(new[]{1,2,3})).ToString());
__Check("empty\nsingle\npair\nmany");
