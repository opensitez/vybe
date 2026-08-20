// Tic Tac Toe with Minimax AI — you are X, computer is O.
//
// Built on the WHATWG DOM: the nine cells are `<button>` elements held in an
// array, so a move writes to the element it already has. The old version
// addressed them by a made-up name ("cell" + i) and asked a toolkit to look
// each one up; an element handle IS the identity, so no lookup happens here.
document.setTitle("Tic Tac Toe");

let board = ["", "", "", "", "", "", "", "", ""];
let gameOver = false;

let status = document.createElement("div");
status.setTextContent("Your turn (X)");
document.body.appendChild(status);

// The nine cell elements, in board order.
let cells = [];

function setCell(idx, mark) {
    board[idx] = mark;
    cells[idx].setTextContent(mark);
}

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
                setCell(i, "O");
                return;
            }
            board[i] = "";
        }
    }
    for (let i = 0; i < 9; i++) {
        if (board[i] === "") {
            board[i] = "X";
            if (checkWinner(board) === "X") {
                setCell(i, "O");
                return;
            }
            board[i] = "";
        }
    }

    // Take center if available
    if (board[4] === "") {
        setCell(4, "O");
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
        setCell(bestIdx, "O");
    }
}

function updateBoard() {
    for (let i = 0; i < 9; i++) {
        cells[i].setTextContent(board[i]);
    }
}

function makeCellHandler(idx) {
    return () => {
        if (gameOver) { return; }
        if (board[idx] !== "") { return; }

        // Human plays X
        setCell(idx, "X");

        let w = checkWinner(board);
        if (w === "X") {
            status.setTextContent("You win!");
            gameOver = true;
            return;
        }
        if (isFull(board)) {
            status.setTextContent("Draw!");
            gameOver = true;
            return;
        }

        // Computer plays O
        status.setTextContent("Thinking...");
        computerMove();

        w = checkWinner(board);
        if (w === "O") {
            status.setTextContent("Computer wins!");
            gameOver = true;
            return;
        }
        if (isFull(board)) {
            status.setTextContent("Draw!");
            gameOver = true;
            return;
        }

        status.setTextContent("Your turn (X)");
    };
}

// Three cells per row. A row is a `<div>` and the cells inside it are inline,
// so the document lays the grid out — nothing computes an x or a y.
let row = null;
for (let i = 0; i < 9; i++) {
    if (i % 3 === 0) {
        row = document.createElement("div");
        document.body.appendChild(row);
    }
    let cell = document.createElement("button");
    row.appendChild(cell);
    cells[i] = cell;
    cell.addEventListener("click", makeCellHandler(i));
}

let reset = document.createElement("button");
reset.setTextContent("New Game");
document.body.appendChild(reset);
reset.addEventListener("click", () => {
    board = ["", "", "", "", "", "", "", "", ""];
    gameOver = false;
    status.setTextContent("Your turn (X)");
    updateBoard();
});
