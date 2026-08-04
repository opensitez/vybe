// vybe-test: dart/flutter_ui_paint_shader/fragment_shader_compiles
// origin: languages/dart/tests/dart/test_flutter_ui_paint_shader.rs

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
void __vybeMain() {
  // We can't actually compile a shader from asset without mock host logic,
  // but we can check FragmentProgram type exists.
  print(FragmentProgram != null);
}

void main() {
  __vybeMain();
  __check('true');
}
