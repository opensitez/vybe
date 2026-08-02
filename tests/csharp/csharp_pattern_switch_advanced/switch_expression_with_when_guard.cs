// vybe-test: csharp/csharp_pattern_switch_advanced/switch_expression_with_when_guard
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Classify(int n)=>n switch{
    int x when x<0=>"negative",
    0=>"zero",
    int x when x%2==0=>"even",
    _=>"odd"};
__Check((Classify(-5)).ToString(), "negative");
__Check((Classify(0)).ToString(), "zero");
__Check((Classify(4)).ToString(), "even");
__Check((Classify(7)).ToString(), "odd");
