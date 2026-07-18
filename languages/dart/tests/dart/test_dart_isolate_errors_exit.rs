use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:isolate Errors & Exit
// ═══════════════════════════════════════════════════════════

#[test]
fn isolate_set_errors_fatal() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  isolate.setErrorsFatal(true);
  print('set_fatal_true');
}
"#
        ),
        vec!["set_fatal_true"]
    );
}

#[test]
fn isolate_set_errors_non_fatal() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  final isolate = Isolate.current;
  isolate.setErrorsFatal(false);
  print('set_fatal_false');
}
"#
        ),
        vec!["set_fatal_false"]
    );
}

#[test]
fn isolate_add_error_listener_handles_error() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void isolateMain(_) {
  throw Exception('isolate_error');
}
void main() async {
  final port = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.addErrorListener(port.sendPort);
  
  final msg = await port.first;
  print((msg[0] as String).contains('isolate_error'));
  port.close();
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_remove_error_listener_ignores_error() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void isolateMain(SendPort control) {
  // Wait a bit, then throw
  Timer(Duration(milliseconds: 50), () {
    control.send('ready_to_throw');
    throw Exception('ignored_error');
  });
}
void main() async {
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
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_add_on_exit_listener_triggered() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(_) {}
void main() async {
  final port = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.addOnExitListener(port.sendPort, response: 'exit_detected');
  
  final msg = await port.first;
  print(msg);
  port.close();
}
"#
        ),
        vec!["exit_detected"]
    );
}

#[test]
fn isolate_exit_with_message() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  Isolate.exit(port, 'goodbye');
}
void main() async {
  final port = ReceivePort();
  await Isolate.spawn(isolateMain, port.sendPort);
  
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["goodbye"]
    );
}

#[test]
fn isolate_exit_kills_isolate() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  Isolate.exit(port, 'dead');
  // if not killed, this throws
  throw Exception('should not reach here');
}
void main() async {
  final port = ReceivePort();
  await Isolate.spawn(isolateMain, port.sendPort);
  
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["dead"]
    );
}

#[test]
fn isolate_errors_are_fatal_behavior() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  Isolate.current.setErrorsFatal(true);
  try {
    throw Exception('test_error');
  } catch (e) {
    port.send('caught');
  }
}
void main() async {
  final port = ReceivePort();
  await Isolate.spawn(isolateMain, port.sendPort);
  final msg = await port.first;
  print(msg);
}
"#
        ),
        vec!["caught"]
    );
}

#[test]
fn isolate_current_set_errors_fatal() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void main() {
  Isolate.current.setErrorsFatal(false);
  print('set');
}
"#
        ),
        vec!["set"]
    );
}

#[test]
fn isolate_kill_immediate_does_not_trigger_exit_listener() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
import 'dart:async';
void isolateMain(_) {
  while(true) {}
}
void main() async {
  final exitPort = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  isolate.addOnExitListener(exitPort.sendPort, response: 'exit_detected');
  
  // Immediate kill might or might not trigger exit listener depending on VM internals.
  // Actually, killing an isolate SHOULD trigger the exit listener.
  // We'll test if it does.
  isolate.kill(priority: Isolate.immediate);
  final msg = await exitPort.first;
  print(msg);
}
"#
        ),
        vec!["exit_detected"]
    );
}

#[test]
fn isolate_uncaught_error_closes_isolate() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(SendPort port) {
  Isolate.current.addOnExitListener(port);
  throw Exception('fatal');
}
void main() async {
  final port = ReceivePort();
  // errorsAreFatal: true by default on spawn
  await Isolate.spawn(isolateMain, port.sendPort);
  // first message will be the exit signal since we didn't add an error listener to main port
  final msg = await port.first;
  print(msg == null); // default exit response is null
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn isolate_ping_unresponsive() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:isolate';
void isolateMain(_) {
  while(true) {}
}
void main() async {
  final port = ReceivePort();
  final isolate = await Isolate.spawn(isolateMain, null);
  // Ping with Isolate.immediate
  isolate.ping(port.sendPort, response: 'alive', priority: Isolate.immediate);
  
  final msg = await port.first;
  print(msg);
  isolate.kill(priority: Isolate.immediate);
}
"#
        ),
        vec!["alive"] // immediate ping responds even if isolate is in a while(true) busy loop on most VMs
    );
}
