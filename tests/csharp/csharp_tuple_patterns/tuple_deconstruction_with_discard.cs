// vybe-test: csharp/csharp_tuple_patterns/tuple_deconstruction_with_discard
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(string name,_,int score)=("Bob",99,"skip",55) switch{
    var t=>(t.Item1,t.Item2,t.Item4)};
__Check((name).ToString(), "Bob"); __Check((score).ToString(), "55");
