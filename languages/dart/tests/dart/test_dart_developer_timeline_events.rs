use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:developer Timeline Events
// ═══════════════════════════════════════════════════════════

#[test]
fn timeline_start_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  Timeline.startSync('myTask');
  Timeline.finishSync();
  print('timeline_sync_done');
}
"#
        ),
        vec!["timeline_sync_done"]
    );
}

#[test]
fn timeline_time_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final result = Timeline.timeSync('myTask2', () {
    return 42;
  });
  print(result);
}
"#
        ),
        vec!["42"]
    );
}

#[test]
fn timeline_instant_sync() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  Timeline.instantSync('instantEvent', arguments: {'count': 10});
  print('instant_done');
}
"#
        ),
        vec!["instant_done"]
    );
}

#[test]
fn timeline_task_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final task = TimelineTask();
  task.start('asyncTask');
  task.pass(); // pass down to nested
  task.finish();
  print('task_done');
}
"#
        ),
        vec!["task_done"]
    );
}

#[test]
fn timeline_task_with_id() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final task = TimelineTask.withTaskId(12345);
  task.instant('checkpoint');
  task.finish();
  print('task_with_id_done');
}
"#
        ),
        vec!["task_with_id_done"]
    );
}

#[test]
fn flow_creation() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final flow = Flow.begin();
  print(flow.id > 0);
  Flow.step(flow.id);
  Flow.end(flow.id);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn timeline_now() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  final now = Timeline.now;
  // It returns microseconds since some epoch
  print(now > 0);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn timeline_time_sync_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:developer';
void main() {
  try {
    Timeline.timeSync('throwingTask', () {
      throw StateError('abort');
    });
  } catch(e) {
    print('caught');
  }
}
"#
        ),
        vec!["caught"] // The timeline event should be closed correctly even if throwing
    );
}
