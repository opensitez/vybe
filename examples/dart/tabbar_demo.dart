import 'package:flutter/material.dart';

void main() {
  runApp(const TabDemo());
}

class TabDemo extends StatelessWidget {
  const TabDemo();
  @override
  Widget build(BuildContext context) {
    return TabBar(
      tabs: [
        Tab(text: "Home"),
        Tab(text: "Search"),
        Tab(text: "Profile"),
      ],
    );
  }
}
