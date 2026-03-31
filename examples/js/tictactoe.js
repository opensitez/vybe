// Tic Tac Toe with Minimax AI — you are X, computer is O
let form = gui.createForm("Tic Tac Toe");
gui.setProperty(form, "Width", 340);
gui.setProperty(form, "Height", 430);

let board = ["", "", "", "", "", "", "", "", ""];
let gameOver = false;

gui.addControl(form, "Label", "status", 10, 10, 310, 30);
gui.setProperty("status", "Text", "Your turn (X)");

function checkWinner(b) {
    let lines = [
        [0,1,2],[3,4,5],[6,7,8],
        [0,3,6],[1,4,7],[2,5,8],
        [0,4,8],[2,4,6]
    ];
    for (let i = 0; i < 8; i++) {
        let a = lines[i][0];
        let c = lines[i][1];
        let d = lines[i][2];
        if (b[a] !== "" && b[a] === b[c] && b[a] === b[d]) {
            return b[a];
        }
    }
    return "";
}

function isFull(b) {
    for (let i = 0; i < 9; i++) {
        if (b[i] === "") { return false; }
    }
    return true;
}

function minimax(b, isMaximizing, depth) {
    let winner = checkWinner(b);
    if (winner === "O") { return 10 - depth; }
    if (winner === "X") { return depth - 10; }
    if (isFull(b)) { return 0; }
    if (depth >= 4) { return 0; }

    if (isMaximizing) {
        let best = -100;
        for (let i = 0; i < 9; i++) {
            if (b[i] === "") {
                b[i] = "O";
                let score = minimax(b, false, depth + 1);
                b[i] = "";
                if (score > best) { best = score; }
            }
        }
        return best;
    } else {
        let best = 100;
        for (let i = 0; i < 9; i++) {
            if (b[i] === "") {
                b[i] = "X";
                let score = minimax(b, true, depth + 1);
                b[i] = "";
                if (score < best) { best = score; }
            }
        }
        return best;
    }
}

function computerMove() {
    // Try to win or block immediately
    for (let i = 0; i < 9; i++) {
        if (board[i] === "") {
            board[i] = "O";
            if (checkWinner(board) === "O") {
                gui.setProperty("cell" + i, "Text", "O");
                return;
            }
            board[i] = "";
        }
    }
    for (let i = 0; i < 9; i++) {
        if (board[i] === "") {
            board[i] = "X";
            if (checkWinner(board) === "X") {
                board[i] = "O";
                gui.setProperty("cell" + i, "Text", "O");
                return;
            }
            board[i] = "";
        }
    }

    // Take center if available
    if (board[4] === "") {
        board[4] = "O";
        gui.setProperty("cell4", "Text", "O");
        return;
    }

    // Use limited minimax for the rest
    let bestScore = -100;
    let bestIdx = -1;
    for (let i = 0; i < 9; i++) {
        if (board[i] === "") {
            board[i] = "O";
            let score = minimax(board, false, 0);
            board[i] = "";
            if (score > bestScore) {
                bestScore = score;
                bestIdx = i;
            }
        }
    }
    if (bestIdx >= 0) {
        board[bestIdx] = "O";
        gui.setProperty("cell" + bestIdx, "Text", "O");
    }
}

function updateBoard() {
    for (let i = 0; i < 9; i++) {
        gui.setProperty("cell" + i, "Text", board[i]);
    }
}

function makeCellHandler(idx) {
    return () => {
        if (gameOver) { return; }
        if (board[idx] !== "") { return; }

        // Human plays X
        board[idx] = "X";
        gui.setProperty("cell" + idx, "Text", "X");

        let w = checkWinner(board);
        if (w === "X") {
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
        computerMove();

        w = checkWinner(board);
        if (w === "O") {
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

for (let i = 0; i < 9; i++) {
    let row = Math.floor(i / 3);
    let col = i - row * 3;
    let x = 10 + col * 105;
    let y = 50 + row * 105;
    gui.addControl(form, "Button", "cell" + i, x, y, 100, 100);
    gui.setProperty("cell" + i, "Text", "");
    gui.onEvent("cell" + i, "Click", makeCellHandler(i));
}

gui.addControl(form, "Button", "resetBtn", 10, 370, 310, 35);
gui.setProperty("resetBtn", "Text", "New Game");
gui.onEvent("resetBtn", "Click", () => {
    board = ["", "", "", "", "", "", "", "", ""];
    gameOver = false;
    gui.setProperty("status", "Text", "Your turn (X)");
    updateBoard();
});

gui.runApplication(form);
