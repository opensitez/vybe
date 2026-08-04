// vybe-test: dart/flutter_widgets_stream_builder/stream_builder_creation
// origin: languages/dart/tests/dart/test_flutter_widgets_stream_builder.rs

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

import 'package:flutter/widgets.dart';
void __vybeMain() {
  final sb = StreamBuilder<String>(
    stream: Stream.value('Hello'),
    builder: (BuildContext context, AsyncSnapshot<String> snapshot) {
      return const SizedBox();
    },
  );
  __p(sb is StatefulWidget);
}

void main() {
  __vybeMain();
  __check('true');
}
