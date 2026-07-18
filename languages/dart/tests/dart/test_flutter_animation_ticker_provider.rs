use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Flutter: animation Ticker & TickerProvider
// ═══════════════════════════════════════════════════════════

#[test]
fn ticker_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  print(ticker != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ticker_start() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  ticker.start();
  print(ticker.isActive);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ticker_stop() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  ticker.start();
  ticker.stop();
  print(ticker.isActive);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn ticker_dispose() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  ticker.dispose();
  try {
    ticker.start();
  } catch(e) {
    print('FlutterError');
  }
}
"#
        ),
        vec!["FlutterError"]
    );
}

#[test]
fn ticker_is_ticking_initially_false() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  print(ticker.isTicking);
}
"#
        ),
        vec!["false"]
    );
}

#[test]
fn ticker_muted() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  final ticker = Ticker((_) {});
  ticker.muted = true;
  ticker.start();
  // If muted, it is active but usually not ticking
  print(ticker.isActive);
  print(ticker.isTicking);
}
"#
        ),
        vec!["true\nfalse"]
    );
}

#[test]
fn ticker_future() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() async {
  final ticker = Ticker((_) {});
  final future = ticker.start();
  ticker.stop(canceled: true);
  try {
    await future;
    print('finished');
  } catch(e) {
    print('canceled'); // TickerCanceled
  }
}
"#
        ),
        vec!["canceled"]
    );
}

#[test]
fn ticker_future_complete() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() async {
  final ticker = Ticker((_) {});
  final future = ticker.start();
  ticker.stop(); // canceled: false by default
  await future;
  print('stopped cleanly');
}
"#
        ),
        vec!["stopped cleanly"]
    );
}

#[test]
fn ticker_provider_interface() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
class MyProvider implements TickerProvider {
  @override
  Ticker createTicker(TickerCallback onTick) => Ticker(onTick);
}
void main() {
  final p = MyProvider();
  final t = p.createTicker((_) {});
  print(t != null);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ticker_elapsed_time() {
    assert_eq!(
        run_prints(
            r#"
import 'package:flutter/scheduler.dart';
void main() {
  int count = 0;
  final ticker = Ticker((elapsed) {
    if (elapsed.inMilliseconds >= 0) {
      count++;
    }
  });
  ticker.start();
  // We can't actually wait for ticks natively without scheduler mock,
  // so we just ensure it doesn't crash on start.
  print('started');
  ticker.stop();
}
"#
        ),
        vec!["started"]
    );
}
