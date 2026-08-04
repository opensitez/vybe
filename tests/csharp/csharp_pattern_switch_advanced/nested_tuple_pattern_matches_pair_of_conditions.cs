// vybe-test: csharp/csharp_pattern_switch_advanced/nested_tuple_pattern_matches_pair_of_conditions
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

string Combo(bool a,bool b)=>(a,b) switch{
    (true,true)=>"both",
    (true,false)=>"left",
    (false,true)=>"right",
    _=>"none"};
__P((Combo(true,false)).ToString());
__Check("left");
