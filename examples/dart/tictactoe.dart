// Tic Tac Toe with Minimax AI for Vybe Dart
void main() {
    var form = gui.createForm("Tic Tac Toe");
    gui.setProperty(form, "Width", 340);
    gui.setProperty(form, "Height", 430);

    var board = ["", "", "", "", "", "", "", "", ""];
    var gameOver = false;

    gui.addControl(form, "Label", "status", 10, 10, 310, 30);
    gui.setProperty("status", "Text", "Your turn (X)");

    String checkWinner(b) {
        var lines = [
            [0,1,2],[3,4,5],[6,7,8],
            [0,3,6],[1,4,7],[2,5,8],
            [0,4,8],[2,4,6]
        ];
        for (var i = 0; i < 8; i++) {
            var a = lines[i][0];
            var c = lines[i][1];
            var d = lines[i][2];
            if (b[a] != "" && b[a] == b[c] && b[a] == b[d]) {
                return b[a];
            }
        }
        return "";
    }

    bool isFull(b) {
        for (var i = 0; i < 9; i++) {
            if (b[i] == "") { return false; }
        }
        return true;
    }

    int minimax(b, bool isMaximizing, int depth) {
        var winner = checkWinner(b);
        if (winner == "O") { return 10 - depth; }
        if (winner == "X") { return depth - 10; }
        if (isFull(b)) { return 0; }
        if (depth >= 4) { return 0; }

        if (isMaximizing) {
            var best = -100;
            for (var i = 0; i < 9; i++) {
                if (b[i] == "") {
                    b[i] = "O";
                    var score = minimax(b, false, depth + 1);
                    b[i] = "";
                    if (score > best) { best = score; }
                }
            }
            return best;
        } else {
            var best = 100;
            for (var i = 0; i < 9; i++) {
                if (b[i] == "") {
                    b[i] = "X";
                    var score = minimax(b, true, depth + 1);
                    b[i] = "";
                    if (score < best) { best = score; }
                }
            }
            return best;
        }
    }

    void updateBoard() {
        for (var i = 0; i < 9; i++) {
            gui.setProperty("cell$i", "Text", board[i]);
        }
    }

    void computerMove() {
        // Try to win or block immediately
        for (var i = 0; i < 9; i++) {
            if (board[i] == "") {
                board[i] = "O";
                if (checkWinner(board) == "O") {
                    gui.setProperty("cell$i", "Text", "O");
                    return;
                }
                board[i] = "";
            }
        }
        for (var i = 0; i < 9; i++) {
            if (board[i] == "") {
                board[i] = "X";
                if (checkWinner(board) == "X") {
                    board[i] = "O";
                    gui.setProperty("cell$i", "Text", "O");
                    return;
                }
                board[i] = "";
            }
        }

        // Take center if available
        if (board[4] == "") {
            board[4] = "O";
            gui.setProperty("cell4", "Text", "O");
            return;
        }

        // Use limited minimax for the rest
        var bestScore = -100;
        var bestIdx = -1;
        for (var i = 0; i < 9; i++) {
            if (board[i] == "") {
                board[i] = "O";
                var score = minimax(board, false, 0);
                board[i] = "";
                if (score > bestScore) {
                    bestScore = score;
                    bestIdx = i;
                }
            }
        }
        if (bestIdx >= 0) {
            board[bestIdx] = "O";
            gui.setProperty("cell$bestIdx", "Text", "O");
        }
    }

    Function makeCellHandler(int idx) {
        return () {
            if (gameOver) { return; }
            if (board[idx] != "") { return; }

            // Human plays X
            board[idx] = "X";
            gui.setProperty("cell$idx", "Text", "X");

            var w = checkWinner(board);
            if (w == "X") {
                gui.setProperty("status", "Text", "You win!");
                gameOver = true;
                return;
            }
            if (isFull(board)) {
                gui.setProperty("status", "Text", "Draw!");
                gameOver = true;
                return;
            }

            // Computer plays O
            gui.setProperty("status", "Text", "Thinking...");
            // Minimal delay usually expected here, but we compute immediately
            computerMove();

            w = checkWinner(board);
            if (w == "O") {
                gui.setProperty("status", "Text", "Computer wins!");
                gameOver = true;
                return;
            }
            if (isFull(board)) {
                gui.setProperty("status", "Text", "Draw!");
                gameOver = true;
                return;
            }

            gui.setProperty("status", "Text", "Your turn (X)");
        };
    }

    // Initialize 3x3 grid
    for (var i = 0; i < 9; i++) {
        var col = i % 3;
        var row = (i - col) / 3;
        var x = 10 + col * 105;
        var y = 50 + row * 105;
        gui.addControl(form, "Button", "cell$i", x, y, 100, 100);
        gui.setProperty("cell$i", "Text", "");
        gui.onEvent("cell$i", "Click", makeCellHandler(i));
    }

    gui.addControl(form, "Button", "resetBtn", 10, 370, 310, 35);
    gui.setProperty("resetBtn", "Text", "New Game");
    gui.onEvent("resetBtn", "Click", () {
        board = ["", "", "", "", "", "", "", "", ""];
        gameOver = false;
        gui.setProperty("status", "Text", "Your turn (X)");
        updateBoard();
    });

    gui.runApplication(form);
}
