import 'package:flutter/material.dart';

void main() {
  runApp(const DropdownDemo());
}

class DropdownDemo extends StatefulWidget {
  const DropdownDemo();
  @override
  State<DropdownDemo> createState() => _DropdownDemoState();
}

class _DropdownDemoState extends State<DropdownDemo> {
  String _fruit = "Banana";
  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        Text("Picked: $_fruit"),
        DropdownButton(
          value: _fruit,
          items: [
            DropdownMenuItem(value: "Apple", child: Text("Apple")),
            DropdownMenuItem(value: "Banana", child: Text("Banana")),
            DropdownMenuItem(value: "Cherry", child: Text("Cherry")),
          ],
          onChanged: (String v) {
            setState(() {
              _fruit = v;
            });
          },
        ),
      ],
    );
  }
}
