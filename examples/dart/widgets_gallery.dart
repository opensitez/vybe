// Widget gallery — exercises the Flutter→vybe_widgets adapter across the
// interactive control family (Checkbox, Slider, TextField, Radio, Dropdown,
// progress, buttons). Each Flutter widget maps onto its real vybe:gui control.
import 'package:flutter/material.dart';

void main() {
  runApp(const GalleryApp());
}

class GalleryApp extends StatelessWidget {
  const GalleryApp();
  @override
  Widget build(BuildContext context) {
    return MaterialApp(home: GalleryPage());
  }
}

class GalleryPage extends StatefulWidget {
  @override
  State<GalleryPage> createState() => _GalleryState();
}

class _GalleryState extends State<GalleryPage> {
  bool checked = true;
  double slider = 40.0;
  int radioGroup = 1;
  String field = "hello";

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text("Widget Gallery")),
      body: Column(
        children: [
          Text("Interactive controls over vybe_widgets"),
          Checkbox(
            value: checked,
            onChanged: (bool? v) {
              setState(() {
                checked = v == true;
              });
            },
          ),
          Slider(
            value: slider,
            min: 0.0,
            max: 100.0,
            onChanged: (double v) {
              setState(() {
                slider = v;
              });
            },
          ),
          TextField(
            controller: TextEditingController(text: field),
            onChanged: (String v) {
              setState(() {
                field = v;
              });
            },
          ),
          Radio<int>(
            value: 1,
            groupValue: radioGroup,
            onChanged: (int? v) {
              setState(() {
                radioGroup = v ?? 1;
              });
            },
          ),
          LinearProgressIndicator(value: 0.6),
          ElevatedButton(
            onPressed: () {
              setState(() {
                slider = 80.0;
              });
            },
            child: Text("Boost slider"),
          ),
        ],
      ),
    );
  }
}
