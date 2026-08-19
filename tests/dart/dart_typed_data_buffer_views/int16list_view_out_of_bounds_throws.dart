// vybe-test: dart/dart_typed_data_buffer_views/int16list_view_out_of_bounds_throws
// origin: languages/dart/tests/dart/test_dart_typed_data_buffer_views.rs

import 'dart:typed_data';

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

void __vybeMain() {
  final bd = ByteData(3);
  try {
    Int16List.view(bd.buffer, 2);
  } on ArgumentError {
    __p('ArgumentError thrown');
  }
}

void main() {
  __vybeMain();
  __check('ArgumentError thrown');
}
