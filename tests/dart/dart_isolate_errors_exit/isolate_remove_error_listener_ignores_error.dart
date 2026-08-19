// vybe-test: dart/dart_isolate_errors_exit/isolate_remove_error_listener_ignores_error
// origin: languages/dart/tests/dart/test_dart_isolate_errors_exit.rs

import 'dart:isolate';
import 'dart:async';

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

void isolateMain(SendPort control) {
  // Wait a bit, then throw
  Timer(Duration(milliseconds: 50), () {
    control.send('ready_to_throw');
    throw Exception('ignored_error');
  });
}
void __vybeMain() async {
  final port = ReceivePort();
  final controlPort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, controlPort.sendPort);
  
  isolate.addErrorListener(port.sendPort);
  isolate.removeErrorListener(port.sendPort); // immediately remove
  
  // wait for it to be ready
  await controlPort.first;
  controlPort.close();
  
  // if error was caught, it would arrive on `port`. We give it 100ms
  // Actually, dart isolates that throw unhandled might crash the whole app if errorsAreFatal is true
  // Let's set errorsAreFatal to false explicitly
  isolate.setErrorsFatal(false);
  
  var receivedError = false;
  final sub = port.listen((_) { receivedError = true; });
  
  await Future.delayed(Duration(milliseconds: 100));
  print(receivedError == false);
  sub.cancel();
  port.close();
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
