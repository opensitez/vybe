// vybe-test: csharp/csharp_pattern_switch_advanced/relational_and_pattern_combines_bounds_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Grade(int n)=>n switch{
    >=90=>"A",
    >=70 and <90=>"B",
    >=50 and <70=>"C",
    _=>"F"};
__Check((Grade(95)).ToString(), "A");
__Check((Grade(75)).ToString(), "B");
__Check((Grade(55)).ToString(), "C");
__Check((Grade(30)).ToString(), "F");
