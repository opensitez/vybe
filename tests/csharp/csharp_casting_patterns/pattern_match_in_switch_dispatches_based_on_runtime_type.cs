// vybe-test: csharp/csharp_casting_patterns/pattern_match_in_switch_dispatches_based_on_runtime_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=42;
string r=o switch{int n=>$"int:{n}",string s=>$"str:{s}",_=>"other"};
__Check((r).ToString(), "int:42");
