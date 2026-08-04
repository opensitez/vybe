// vybe-test: dart/streams_core/stream_length_of_value_stream_is_one
// origin: languages/dart/tests/dart/test_streams_core.rs

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

Future<void> __vybeMain() async {
  var n = await Stream.value(0).length;
  __p(n);
}

Future<void> main() async {
  await __vybeMain();
  __check('1');
}
