import 'package:flutter/material.dart';

void main() {
  runApp(const SwitchDemo());
}

class SwitchDemo extends StatefulWidget {
  const SwitchDemo();
  @override
  State<SwitchDemo> createState() => _SwitchDemoState();
}

class _SwitchDemoState extends State<SwitchDemo> {
  bool _on = false;
  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        Text(_on ? "ON" : "OFF"),
        Switch(
          value: _on,
          onChanged: (bool v) {
            setState(() {
              _on = v;
            });
          },
        ),
      ],
    );
  }
}
