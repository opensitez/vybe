// vybe-test: csharp/csharp_properties_accessors/expression_bodied_setter_updates_hidden_field
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

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

class ScoreBoard {
    int score;
    public int Score {
        get => score;
        set => score = value * 10;
    }
}
var board = new ScoreBoard();
board.Score = 7;
__P((board.Score).ToString());
__Check("70");
