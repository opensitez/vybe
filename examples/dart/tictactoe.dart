// Tic Tac Toe with a Minimax AI — a Flutter app built from Dart.
//
// Idiomatic Flutter to the source; under the hood the Vybe `flutter` platform
// adapter drives the shared `vybe_widgets` runtime through the same common
// resolver as the WinForms/VCL adapters — no Flutter-specific host functions.
import 'package:flutter/material.dart';

void main() {
  runApp(const TicTacToeApp());
}

class TicTacToeApp extends StatelessWidget {
  const TicTacToeApp();

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Tic Tac Toe',
      home: const GamePage(),
    );
  }
}

class GamePage extends StatefulWidget {
  const GamePage();

  @override
  State<GamePage> createState() => _GamePageState();
}

class _GamePageState extends State<GamePage> {
  List<String> board = ["", "", "", "", "", "", "", "", ""];
  bool gameOver = false;
  String status = "Your turn (X)";

  String checkWinner(List<String> b) {
    final lines = [
      [0, 1, 2], [3, 4, 5], [6, 7, 8],
      [0, 3, 6], [1, 4, 7], [2, 5, 8],
      [0, 4, 8], [2, 4, 6],
    ];
    for (var i = 0; i < 8; i++) {
      final a = lines[i][0];
      final c = lines[i][1];
      final d = lines[i][2];
      if (b[a] != "" && b[a] == b[c] && b[a] == b[d]) {
        return b[a];
      }
    }
    return "";
  }

  bool isFull(List<String> b) {
    for (var i = 0; i < 9; i++) {
      if (b[i] == "") {
        return false;
      }
    }
    return true;
  }

  int minimax(List<String> b, bool isMaximizing, int depth) {
    final winner = checkWinner(b);
    if (winner == "O") {
      return 10 - depth;
    }
    if (winner == "X") {
      return depth - 10;
    }
    if (isFull(b) || depth >= 4) {
      return 0;
    }
    if (isMaximizing) {
      var best = -100;
      for (var i = 0; i < 9; i++) {
        if (b[i] == "") {
          b[i] = "O";
          final score = minimax(b, false, depth + 1);
          b[i] = "";
          if (score > best) {
            best = score;
          }
        }
      }
      return best;
    } else {
      var best = 100;
      for (var i = 0; i < 9; i++) {
        if (b[i] == "") {
          b[i] = "X";
          final score = minimax(b, true, depth + 1);
          b[i] = "";
          if (score < best) {
            best = score;
          }
        }
      }
      return best;
    }
  }

  void computerMove() {
    // Win immediately if possible.
    for (var i = 0; i < 9; i++) {
      if (board[i] == "") {
        board[i] = "O";
        if (checkWinner(board) == "O") {
          return;
        }
        board[i] = "";
      }
    }
    // Block the human's immediate win.
    for (var i = 0; i < 9; i++) {
      if (board[i] == "") {
        board[i] = "X";
        if (checkWinner(board) == "X") {
          board[i] = "O";
          return;
        }
        board[i] = "";
      }
    }
    // Take the center.
    if (board[4] == "") {
      board[4] = "O";
      return;
    }
    // Otherwise use a bounded minimax.
    var bestScore = -100;
    var bestIdx = -1;
    for (var i = 0; i < 9; i++) {
      if (board[i] == "") {
        board[i] = "O";
        final score = minimax(board, false, 0);
        board[i] = "";
        if (score > bestScore) {
          bestScore = score;
          bestIdx = i;
        }
      }
    }
    if (bestIdx >= 0) {
      board[bestIdx] = "O";
    }
  }

  void playCell(int idx) {
    if (gameOver || board[idx] != "") {
      return;
    }
    setState(() {
      board[idx] = "X";
      final w = checkWinner(board);
      if (w == "X") {
        status = "You win!";
        gameOver = true;
        return;
      }
      if (isFull(board)) {
        status = "Draw!";
        gameOver = true;
        return;
      }
      computerMove();
      final w2 = checkWinner(board);
      if (w2 == "O") {
        status = "Computer wins!";
        gameOver = true;
        return;
      }
      if (isFull(board)) {
        status = "Draw!";
        gameOver = true;
        return;
      }
      status = "Your turn (X)";
    });
  }

  void resetGame() {
    setState(() {
      board = ["", "", "", "", "", "", "", "", ""];
      gameOver = false;
      status = "Your turn (X)";
    });
  }

  Widget buildCell(int idx) {
    return Expanded(
      child: Padding(
        padding: EdgeInsets.all(4.0),
        child: ElevatedButton(
          onPressed: () {
            playCell(idx);
          },
          child: Text(board[idx]),
        ),
      ),
    );
  }

  Widget buildRow(int start) {
    return Expanded(
      child: Row(
        children: [
          buildCell(start),
          buildCell(start + 1),
          buildCell(start + 2),
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Tic Tac Toe')),
      body: Column(
        children: [
          Padding(
            padding: EdgeInsets.all(16.0),
            child: Text(status),
          ),
          buildRow(0),
          buildRow(3),
          buildRow(6),
          Padding(
            padding: EdgeInsets.all(8.0),
            child: ElevatedButton(
              onPressed: () {
                resetGame();
              },
              child: const Text("New Game"),
            ),
          ),
        ],
      ),
    );
  }
}
