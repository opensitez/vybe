use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:isolate Ports & Messaging
// ═══════════════════════════════════════════════════════════

#[test]
fn receive_port_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final port = ReceivePort();
  print(port is Stream);
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn receive_port_send_port() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final port = ReceivePort();
  final sendPort = port.sendPort;
  print(sendPort is SendPort);
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn send_port_send() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  port.sendPort.send('ping');
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["ping"]
    );
}

#[test]
fn send_port_send_multiple() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  port.sendPort.send(1);
  port.sendPort.send(2);
  final list = await port.take(2).toList();
  print('${list[0]}:${list[1]}');
}
"#
        ),
        vec!["1:2"]
    );
}

#[test]
fn receive_port_listen() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  port.listen((msg) {
    if (msg == 'close') {
      port.close();
    } else {
      print(msg);
    }
  });
  port.sendPort.send('hello');
  port.sendPort.send('close');
}
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn send_port_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final port = ReceivePort();
  final sp1 = port.sendPort;
  final sp2 = port.sendPort;
  print(sp1 == sp2);
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn send_port_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final port = ReceivePort();
  final sp1 = port.sendPort;
  final sp2 = port.sendPort;
  print(sp1.hashCode == sp2.hashCode);
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn raw_receive_port_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final port = RawReceivePort();
  print(port.sendPort is SendPort);
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn raw_receive_port_handler() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void main() async {
  final completer = Completer();
  final port = RawReceivePort((msg) {
    completer.complete(msg);
  });
  port.sendPort.send('raw_message');
  final msg = await completer.future;
  print(msg);
  port.close();
}
"#
        ),
        vec!["raw_message"]
    );
}

#[test]
fn raw_receive_port_close_removes_handler() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void main() async {
  final port = RawReceivePort((msg) {
    print('received');
  });
  port.close();
  port.sendPort.send('missed');
  await Future.delayed(Duration(milliseconds: 10));
  print('done');
}
"#
        ),
        vec!["done"] // 'received' should not be printed
    );
}

#[test]
fn send_port_send_null() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  port.sendPort.send(null);
  final msg = await port.first;
  print(msg == null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn send_port_send_complex_object() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  final map = {'a': 1, 'b': [2, 3]};
  port.sendPort.send(map);
  final msg = await port.first;
  print(msg['b'][1]);
}
"#
        ),
        vec!["3"]
    );
}

#[test]
fn send_port_send_closure_fails() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  void myFunc() {}
  try {
    port.sendPort.send(myFunc);
    // Might fail depending on Dart version and closure context
    // Actually, Dart allows sending top-level functions over send ports
    final msg = await port.first;
    print(msg is Function);
  } catch(e) {
    print('throws');
  }
}
"#
        ),
        vec!["true"] // Dart allows sending top-level or static functions
    );
}

#[test]
fn receive_port_as_broadcast_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  final broadcast = port.asBroadcastStream();
  port.sendPort.send('broad');
  final msg = await broadcast.first;
  print(msg);
}
"#
        ),
        vec!["broad"]
    );
}

#[test]
fn receive_port_from_raw() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final raw = RawReceivePort();
  final port = ReceivePort.fromRawReceivePort(raw);
  port.sendPort.send('from_raw');
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["from_raw"]
    );
}

#[test]
fn send_port_send_recursive_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  final port = ReceivePort();
  final list = [];
  list.add(list);
  try {
    port.sendPort.send(list);
    final msg = await port.first;
    print(identical(msg, msg[0]));
  } catch(e) {
    print('throws'); // Older dart might throw ArgumentError
  }
}
"#
        ),
        // Wait, Dart's SendPort message serialization actually supports cyclic graphs since a while ago,
        // but `identical` across isolates or ports might not hold depending on whether it's the same isolate.
        // Within same isolate, it actually copies the object graph, so `msg != list`, but `msg[0] == msg`.
        vec!["true"] 
    );
}
