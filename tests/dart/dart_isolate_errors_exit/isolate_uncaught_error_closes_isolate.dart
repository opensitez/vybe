// vybe-test: dart/dart_isolate_errors_exit/isolate_uncaught_error_closes_isolate
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

import 'dart:isolate';

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

void isolateMain(SendPort port) {
  Isolate.current.addOnExitListener(port);
  throw Exception('fatal');
}
void __vybeMain() async {
  final port = ReceivePort();
  // errorsAreFatal: true by default on spawn
  await Isolate.spawn(isolateMain, port.sendPort);
  // first message will be the exit signal since we didn't add an error listener to main port
  final msg = await port.first;
  print(msg == null); // default exit response is null
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
