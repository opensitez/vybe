import 'package:flutter/material.dart';

void main() {
  runApp(const ListDemo());
}

class ListDemo extends StatelessWidget {
  const ListDemo();
  @override
  Widget build(BuildContext context) {
    return ListView(
      children: [
        Text("Alpha"),
        Text("Bravo"),
        Text("Charlie"),
        ListTile(title: Text("Delta")),
      ],
    );
  }
}
