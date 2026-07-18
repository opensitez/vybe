use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:isolate Capabilities & Pause/Kill
// ═══════════════════════════════════════════════════════════

#[test]
fn capability_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final cap = Capability();
  print(cap != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn capability_equality() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final cap1 = Capability();
  final cap2 = Capability();
  print(cap1 != cap2);
  print(cap1 == cap1);
}
"#
        ),
        vec!["true\ntrue"]
    );
}

#[test]
fn capability_hashcode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final cap = Capability();
  print(cap.hashCode == cap.hashCode);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_pause_capability() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  print(isolate.pauseCapability is Capability?);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_terminate_capability() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  print(isolate.terminateCapability is Capability?);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_pause_resume() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() async {
  // We can't pause current isolate and resume it easily without a deadlock,
  // but we can spawn one, pause it, and resume it.
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn((port) {
    (port as SendPort).send('started');
  }, receivePort.sendPort);
  
  final cap = isolate.pause();
  isolate.resume(cap);
  
  final msg = await receivePort.first;
  print(msg);
}
"#
        ),
        vec!["started"]
    );
}

#[test]
fn isolate_pause_custom_capability() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  port.send('started');
}
void main() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, receivePort.sendPort);
  
  final cap = Capability();
  isolate.pause(cap);
  isolate.resume(cap);
  
  final msg = await receivePort.first;
  print(msg);
}
"#
        ),
        vec!["started"]
    );
}

#[test]
fn isolate_kill_immediate() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  // busy loop or just wait
  while(true) {}
}
void main() async {
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.kill(priority: Isolate.immediate);
  print('killed');
}
"#
        ),
        vec!["killed"] // Tests that kill(immediate) doesn't hang the VM
    );
}

#[test]
fn isolate_kill_before_next_event() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(_) {}
void main() async {
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.kill(priority: Isolate.beforeNextEvent);
  print('killed');
}
"#
        ),
        vec!["killed"]
    );
}

#[test]
fn isolate_ping() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(_) {}
void main() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.ping(receivePort.sendPort, response: 'pong');
  
  final msg = await receivePort.first;
  print(msg);
  // Actually, ping response might arrive before or after isolate exits.
  // Because it's an empty isolate, it exits fast, so ping might return the 'pong' or exit event depending.
  // It should be 'pong' though.
}
"#
        ),
        vec!["pong"]
    );
}

#[test]
fn isolate_ping_with_custom_response() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void isolateMain(_) {
  // Keep it alive
  Timer(Duration(hours: 1), () {});
}
void main() async {
  final receivePort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.ping(receivePort.sendPort, response: 42);
  
  final msg = await receivePort.first;
  print(msg);
  isolate.kill();
}
"#
        ),
        vec!["42"]
    );
}

#[test]
fn isolate_remove_error_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  final port = ReceivePort();
  isolate.addErrorListener(port.sendPort);
  isolate.removeErrorListener(port.sendPort);
  port.close();
  print('ok');
}
"#
        ),
        vec!["ok"]
    );
}

#[test]
fn isolate_remove_on_exit_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  final port = ReceivePort();
  isolate.addOnExitListener(port.sendPort);
  isolate.removeOnExitListener(port.sendPort);
  port.close();
  print('ok');
}
"#
        ),
        vec!["ok"]
    );
}

#[test]
fn isolate_control_port() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  print(isolate.controlPort is SendPort);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_add_on_exit_listener_response() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(_) {}
void main() async {
  final isolate = await Isolate.spawn(isolateMain, null);
  final port = ReceivePort();
  isolate.addOnExitListener(port.sendPort, response: 'custom_exit');
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["custom_exit"]
    );
}
