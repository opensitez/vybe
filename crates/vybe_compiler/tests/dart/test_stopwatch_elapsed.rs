//! Dart `Stopwatch`: start/stop/reset, elapsed accessors, lap-style structural checks.

dart_cases! {
    stopwatch_new_is_not_running => {
        r#"void main() {
  var sw = Stopwatch();
  print(sw.isRunning);
}"#,
        ["false"]
    };

    stopwatch_elapsed_zero_before_start => {
        r#"void main() {
  var sw = Stopwatch();
  print(sw.elapsed.inMicroseconds);
  print(sw.elapsedMilliseconds);
}"#,
        ["0", "0"]
    };

    stopwatch_start_sets_is_running => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  print(sw.isRunning);
}"#,
        ["true"]
    };

    stopwatch_stop_clears_is_running => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  print(sw.isRunning);
}"#,
        ["false"]
    };

    stopwatch_reset_zeros_elapsed => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  sw.reset();
  print(sw.elapsed.inMicroseconds);
  print(sw.elapsedMilliseconds);
}"#,
        ["0", "0"]
    };

    stopwatch_reset_sets_not_running => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.reset();
  print(sw.isRunning);
  print(sw.elapsed.inMicroseconds);
}"#,
        ["false", "0"]
    };

    stopwatch_elapsed_non_negative_after_start => {
        r#"void main() {
  var sw = Stopwatch()..start();
  print(sw.elapsed.inMicroseconds >= 0);
  print(sw.isRunning);
}"#,
        ["true", "true"]
    };

    stopwatch_elapsed_milliseconds_non_negative_after_start => {
        r#"void main() {
  var sw = Stopwatch()..start();
  print(sw.elapsedMilliseconds >= 0);
}"#,
        ["true"]
    };

    stopwatch_elapsed_frozen_after_stop => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  var first = sw.elapsedMicroseconds;
  var second = sw.elapsedMicroseconds;
  print(first == second);
  print(sw.isRunning);
}"#,
        ["true", "false"]
    };

    stopwatch_elapsed_microseconds_accessor => {
        r#"void main() {
  var sw = Stopwatch();
  print(sw.elapsedMicroseconds);
  sw.start();
  sw.stop();
  print(sw.elapsedMicroseconds >= 0);
}"#,
        ["0", "true"]
    };

    stopwatch_elapsed_milliseconds_zero_after_reset => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  sw.reset();
  print(sw.elapsedMilliseconds);
  print(sw.elapsedMicroseconds);
}"#,
        ["0", "0"]
    };

    stopwatch_double_stop_stays_not_running => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  sw.stop();
  print(sw.isRunning);
}"#,
        ["false"]
    };

    stopwatch_start_after_stop_resumes_running => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  sw.start();
  print(sw.isRunning);
}"#,
        ["true"]
    };

    stopwatch_elapsed_after_stop_then_start_non_negative => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  sw.start();
  sw.stop();
  print(sw.elapsedMicroseconds >= 0);
  print(sw.isRunning);
}"#,
        ["true", "false"]
    };

    stopwatch_lap_read_while_running_structural => {
        r#"void main() {
  var sw = Stopwatch()..start();
  var lap1Running = sw.isRunning;
  var lap1ElapsedOk = sw.elapsed.inMicroseconds >= 0;
  sw.stop();
  var lap1Stopped = sw.isRunning;
  print(lap1Running);
  print(lap1ElapsedOk);
  print(lap1Stopped);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_lap_elapsed_preserved_after_stop => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  var lapMs = sw.elapsedMilliseconds;
  var lapUs = sw.elapsedMicroseconds;
  print(lapMs >= 0);
  print(lapUs >= 0);
  print(sw.isRunning);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_multiple_reset_cycles_zero_elapsed => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  sw.reset();
  sw.start();
  sw.stop();
  sw.reset();
  print(sw.elapsedMilliseconds);
  print(sw.isRunning);
}"#,
        ["0", "false"]
    };

    stopwatch_start_without_prior_start_is_running => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  print(sw.isRunning);
  print(sw.elapsed.inMicroseconds >= 0);
}"#,
        ["true", "true"]
    };

    stopwatch_elapsed_duration_type_has_microseconds => {
        r#"void main() {
  var sw = Stopwatch();
  print(sw.elapsed.inMicroseconds);
  print(sw.elapsed.inMilliseconds);
}"#,
        ["0", "0"]
    };

    stopwatch_reset_after_never_started => {
        r#"void main() {
  var sw = Stopwatch();
  sw.reset();
  print(sw.elapsedMilliseconds);
  print(sw.isRunning);
}"#,
        ["0", "false"]
    };

    stopwatch_stop_without_start_stays_not_running => {
        r#"void main() {
  var sw = Stopwatch();
  sw.stop();
  print(sw.isRunning);
  print(sw.elapsed.inMicroseconds);
}"#,
        ["false", "0"]
    };

    stopwatch_lap_style_start_stop_read_reset => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  print(sw.isRunning);
  sw.stop();
  var lap = sw.elapsedMilliseconds;
  print(lap >= 0);
  sw.reset();
  print(sw.elapsedMilliseconds);
  print(sw.isRunning);
}"#,
        ["true", "true", "0", "false"]
    };

    stopwatch_lap_style_two_segments_structural => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  var seg1 = sw.elapsedMicroseconds;
  sw.start();
  sw.stop();
  var seg2 = sw.elapsedMicroseconds;
  print(seg1 >= 0);
  print(seg2 >= seg1);
  print(sw.isRunning);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_lap_style_reset_between_segments => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  sw.reset();
  sw.start();
  sw.stop();
  print(sw.elapsedMicroseconds >= 0);
  print(sw.elapsedMilliseconds >= 0);
  print(sw.isRunning);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_cascade_start_sets_running => {
        r#"void main() {
  var sw = Stopwatch()..start();
  print(sw.isRunning);
}"#,
        ["true"]
    };

    stopwatch_cascade_start_stop_not_running => {
        r#"void main() {
  var sw = Stopwatch()
    ..start()
    ..stop();
  print(sw.isRunning);
  print(sw.elapsedMicroseconds >= 0);
}"#,
        ["false", "true"]
    };

    stopwatch_cascade_start_stop_reset_zero => {
        r#"void main() {
  var sw = Stopwatch()
    ..start()
    ..stop()
    ..reset();
  print(sw.elapsedMilliseconds);
  print(sw.isRunning);
}"#,
        ["0", "false"]
    };

    stopwatch_elapsed_in_milliseconds_matches_getter => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  print(sw.elapsed.inMilliseconds == sw.elapsedMilliseconds);
}"#,
        ["true"]
    };

    stopwatch_elapsed_in_microseconds_matches_getter => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  print(sw.elapsed.inMicroseconds == sw.elapsedMicroseconds);
}"#,
        ["true"]
    };

    stopwatch_three_lap_reads_all_non_negative => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  var a = sw.elapsedMicroseconds;
  sw.stop();
  var b = sw.elapsedMicroseconds;
  sw.start();
  sw.stop();
  var c = sw.elapsedMicroseconds;
  print(a >= 0);
  print(b >= 0);
  print(c >= 0);
}"#,
        ["true", "true", "true"]
    };

    stopwatch_reset_clears_accumulated_elapsed => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  sw.reset();
  print(sw.elapsedMicroseconds);
  print(sw.elapsed.inMicroseconds);
  print(sw.elapsedMilliseconds);
}"#,
        ["0", "0", "0"]
    };

    stopwatch_is_running_false_after_reset_from_running => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.reset();
  print(sw.isRunning);
}"#,
        ["false"]
    };

    stopwatch_lap_elapsed_milliseconds_after_stop_structural => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  print(sw.elapsedMilliseconds >= 0);
  print(sw.elapsed.inMilliseconds >= 0);
  print(sw.isRunning);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_start_stop_start_stop_structural => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  sw.start();
  sw.stop();
  print(sw.isRunning);
  print(sw.elapsedMicroseconds >= 0);
}"#,
        ["false", "true"]
    };

    stopwatch_elapsed_zero_immediately_after_reset_from_active => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.reset();
  print(sw.elapsed.inMicroseconds);
  print(sw.elapsedMilliseconds);
  print(sw.isRunning);
}"#,
        ["0", "0", "false"]
    };

    stopwatch_lap_compare_before_and_after_stop_same => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  var ms1 = sw.elapsedMilliseconds;
  var ms2 = sw.elapsedMilliseconds;
  var us1 = sw.elapsedMicroseconds;
  var us2 = sw.elapsedMicroseconds;
  print(ms1 == ms2);
  print(us1 == us2);
}"#,
        ["true", "true"]
    };

    stopwatch_multiple_start_while_running_stays_running => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.start();
  print(sw.isRunning);
}"#,
        ["true"]
    };

    stopwatch_elapsed_duration_zero_constant => {
        r#"void main() {
  var sw = Stopwatch();
  print(sw.elapsed == Duration.zero);
  print(sw.elapsed.inMicroseconds);
}"#,
        ["true", "0"]
    };

    stopwatch_lap_sequence_with_interleaved_reset => {
        r#"void main() {
  var sw = Stopwatch();
  sw.start();
  sw.stop();
  var lap1 = sw.elapsedMicroseconds;
  sw.reset();
  sw.start();
  sw.stop();
  var lap2 = sw.elapsedMicroseconds;
  print(lap1 >= 0);
  print(lap2 >= 0);
  print(sw.isRunning);
}"#,
        ["true", "true", "false"]
    };

    stopwatch_stop_preserves_elapsed_until_reset => {
        r#"void main() {
  var sw = Stopwatch()..start();
  sw.stop();
  var before = sw.elapsedMicroseconds;
  var after = sw.elapsedMicroseconds;
  sw.reset();
  print(before == after);
  print(sw.elapsedMicroseconds);
}"#,
        ["true", "0"]
    };
}
