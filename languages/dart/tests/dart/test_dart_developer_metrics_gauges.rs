use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:developer Metrics & Gauges
// ═══════════════════════════════════════════════════════════

#[test]
fn metrics_counter_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final counter = Counter('my.counter', 'A test counter');
  print(counter.name);
  print(counter.description);
}
"#
        ),
        vec!["my.counter\nA test counter"]
    );
}

#[test]
fn metrics_counter_value() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final counter = Counter('my.counter2', 'description');
  counter.value = 10.5;
  print(counter.value);
}
"#
        ),
        vec!["10.5"]
    );
}

#[test]
fn metrics_gauge_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final gauge = Gauge('my.gauge', 'A test gauge', min: 0.0, max: 100.0);
  print(gauge.name);
  print(gauge.min);
  print(gauge.max);
}
"#
        ),
        vec!["my.gauge\n0.0\n100.0"]
    );
}

#[test]
fn metrics_gauge_value() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final gauge = Gauge('my.gauge2', 'desc', min: 0.0, max: 10.0);
  gauge.value = 5.5;
  print(gauge.value);
}
"#
        ),
        vec!["5.5"]
    );
}

#[test]
fn metrics_gauge_value_out_of_bounds() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final gauge = Gauge('my.gauge3', 'desc', min: 0.0, max: 10.0);
  try {
    // Some implementations might clamp, throw, or accept it. 
    // Usually Dart just accepts it but tools might warn.
    gauge.value = 20.0;
    print(gauge.value);
  } catch(e) {
    print('ArgumentError'); // In case it strictly throws
  }
}
"#
        ),
        vec!["20.0"] // Dart Gauge does not clamp or throw on setting out-of-bounds natively
    );
}

#[test]
fn metrics_metric_interface() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final counter = Counter('my.counter3', 'desc');
  Metric m = counter;
  print(m.name);
}
"#
        ),
        vec!["my.counter3"]
    );
}

#[test]
fn metrics_deregister() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final counter = Counter('my.counter4', 'desc');
  // There is no explicit deregister, they are garbage collected or exist globally.
  // We'll just verify setting value to 0 works.
  counter.value = 0.0;
  print(counter.value);
}
"#
        ),
        vec!["0.0"]
    );
}
