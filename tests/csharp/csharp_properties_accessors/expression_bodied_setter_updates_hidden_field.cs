// vybe-test: csharp/csharp_properties_accessors/expression_bodied_setter_updates_hidden_field
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var board = new ScoreBoard();
board.Score = 7;
__P((board.Score).ToString());
__Check("70");

class ScoreBoard {
    int score;
    public int Score {
        get => score;
        set => score = value * 10;
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
