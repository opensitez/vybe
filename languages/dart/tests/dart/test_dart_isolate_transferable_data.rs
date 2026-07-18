use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:isolate Transferable Data
// ═══════════════════════════════════════════════════════════

#[test]
fn transferable_typed_data_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([1, 2, 3]);
  final ttd = TransferableTypedData.fromList([list]);
  print(ttd != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn transferable_typed_data_materialize() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([10, 20, 30]);
  final ttd = TransferableTypedData.fromList([list]);
  final materialized = ttd.materialize();
  print('${materialized.lengthInBytes}:${materialized.getUint8(1)}');
}
"#
        ),
        vec!["3\n20"]
    );
}

#[test]
fn transferable_typed_data_transfers_ownership() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([10, 20, 30]);
  final ttd = TransferableTypedData.fromList([list]);
  // After transfer, 'list' might become inaccessible or cleared depending on Dart version.
  // In Dart >= 2.15, the original list's buffer is detached.
  try {
    list[0];
    print('accessible');
  } catch(e) {
    print('detached');
  }
}
"#
        ),
        vec!["detached"] // Usually accessing detached buffer throws StateError or similar
    );
}

#[test]
fn transferable_typed_data_materialize_multiple_times() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final list = Uint8List.fromList([5, 6, 7]);
  final ttd = TransferableTypedData.fromList([list]);
  final m1 = ttd.materialize();
  final m2 = ttd.materialize();
  // Multiple materialize calls usually return the exact same ByteData instance or a new one pointing to same memory.
  // But actually the memory is moved back to the current isolate context.
  print(m1.lengthInBytes == m2.lengthInBytes);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_send_transferable_typed_data() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void isolateMain(SendPort port) {
  final list = Uint8List.fromList([99, 100]);
  final ttd = TransferableTypedData.fromList([list]);
  port.send(ttd);
}
void main() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final msg = await receivePort.first;
  final bd = (msg as TransferableTypedData).materialize();
  print(bd.getUint8(0));
}
"#
        ),
        vec!["99"]
    );
}

#[test]
fn transferable_typed_data_multiple_lists() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final l1 = Uint8List.fromList([1, 2]);
  final l2 = Uint8List.fromList([3, 4]);
  // TransferableTypedData concatenates the byte data of the lists
  final ttd = TransferableTypedData.fromList([l1, l2]);
  final bd = ttd.materialize();
  print(bd.lengthInBytes);
  print(bd.getUint8(2)); // Should be 3
}
"#
        ),
        vec!["4\n3"]
    );
}

#[test]
fn send_port_send_transfer_capability() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  final cap = Capability();
  port.send(cap);
}
void main() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  final msg = await receivePort.first;
  print(msg is Capability);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn send_port_send_send_port() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  final innerPort = ReceivePort();
  port.send(innerPort.sendPort);
  innerPort.listen((msg) {
    if (msg == 'ping') port.send('pong');
  });
}
void main() async {
  final receivePort = ReceivePort();
  await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final innerSendPort = await receivePort.first as SendPort;
  innerSendPort.send('ping');
  
  // listen for pong on the same receivePort (or a new one if isolate sent it there)
  // Wait, the isolate sent the first msg, and then sends 'pong' to 'port'
  // So receivePort will get another message
  final list = await receivePort.take(2).toList();
  print(list[1]);
}
"#
        ),
        vec!["pong"]
    );
}

#[test]
fn transferable_typed_data_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final ttd = TransferableTypedData.fromList([]);
  final bd = ttd.materialize();
  print(bd.lengthInBytes);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn transferable_typed_data_large() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final l1 = Uint8List(100000);
  final ttd = TransferableTypedData.fromList([l1]);
  final bd = ttd.materialize();
  print(bd.lengthInBytes);
}
"#
        ),
        vec!["100000"]
    );
}

#[test]
fn isolate_send_large_data() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  final list = List.filled(1000, 42); // Send by value / serialized copy
  port.send(list);
}
void main() async {
  final port = ReceivePort();
  await Isolate.spawn(isolateMain, port.sendPort);
  
  final msg = await port.first;
  print((msg as List).length);
}
"#
        ),
        vec!["1000"]
    );
}

#[test]
fn transferable_typed_data_float32list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:typed_data';
void main() {
  final l = Float32List(2);
  l[0] = 3.5;
  final ttd = TransferableTypedData.fromList([l.buffer.asUint8List()]);
  final bd = ttd.materialize();
  print(bd.getFloat32(0, Endian.host));
}
"#
        ),
        vec!["3.5"]
    );
}
