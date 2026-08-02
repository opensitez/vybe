// vybe-test: csharp/csharp_tuple_patterns/tuple_pattern_in_switch_expression_dispatches_on_pair
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Classify(int x,int y)=>(x,y) switch{
    (0,0)=>"origin",
    (>0,0)=>"pos-x",
    (0,>0)=>"pos-y",
    _=>"other"};
__Check((Classify(0,0)).ToString(), "origin");
__Check((Classify(3,0)).ToString(), "pos-x");
__Check((Classify(0,5)).ToString(), "pos-y");
__Check((Classify(1,1)).ToString(), "other");
