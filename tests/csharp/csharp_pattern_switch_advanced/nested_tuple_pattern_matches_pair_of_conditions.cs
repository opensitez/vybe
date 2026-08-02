// vybe-test: csharp/csharp_pattern_switch_advanced/nested_tuple_pattern_matches_pair_of_conditions
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Combo(bool a,bool b)=>(a,b) switch{
    (true,true)=>"both",
    (true,false)=>"left",
    (false,true)=>"right",
    _=>"none"};
__Check((Combo(true,false)).ToString(), "left");
