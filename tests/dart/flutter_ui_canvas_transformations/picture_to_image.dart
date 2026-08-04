// vybe-test: dart/flutter_ui_canvas_transformations/picture_to_image
// origin: languages/dart/tests/dart/test_flutter_ui_canvas_transformations.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'dart:ui';
void __vybeMain() async {
  final recorder = PictureRecorder();
  final canvas = Canvas(recorder);
  canvas.drawRect(Rect.fromLTRB(0, 0, 10, 10), Paint());
  final picture = recorder.endRecording();
  
  // This might be async and requires native engine to execute
  try {
    final image = await picture.toImage(10, 10);
    __p('${image.width}:${image.height}');
  } catch(e) {
    // If running headless or without Skia, it might throw
    __p('mock_ok');
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('10:10');
}
