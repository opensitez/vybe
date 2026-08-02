// vybe-test: csharp/csharp_properties_accessors/expression_bodied_setter_updates_hidden_field
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((board.Score).ToString(), "70");
