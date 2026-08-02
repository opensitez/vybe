// vybe-test: csharp/csharp_pattern_switch_advanced/or_pattern_matches_one_of_several_values
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Weekend(string day)=>day switch{
    "Saturday" or "Sunday"=>"weekend",
    _=>"weekday"};
__Check((Weekend("Saturday")).ToString(), "weekend");
__Check((Weekend("Monday")).ToString(), "weekday");
