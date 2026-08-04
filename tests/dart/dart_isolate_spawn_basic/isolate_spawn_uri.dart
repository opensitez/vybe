// vybe-test: dart/dart_isolate_spawn_basic/isolate_spawn_uri
// origin: languages/dart/tests/dart/test_dart_isolate_spawn_basic.rs

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

import 'dart:isolate';
import 'dart:io';
void __vybeMain() async {
  final dir = Directory.systemTemp.createTempSync('iso_');
  final file = File('${dir.path}/iso.dart');
  file.writeAsStringSync("import 'dart:isolate'; void main(List<String> args, SendPort port) { port.send(args[0]); }");
  final receivePort = ReceivePort();
  try {
    await Isolate.spawnUri(Uri.file(file.path), ['spawn_uri_test'], receivePort.sendPort);
    final msg = await receivePort.first;
    __p(msg);
  } finally {
    dir.deleteSync(recursive: true);
  }
}

Future<void> main() async {
  await __vybeMain();
  __check('spawn_uri_test');
}
