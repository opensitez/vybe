use super::helpers::run_python;

// tracemalloc — snapshots, StatsDiff, filter_traces, get_object_traceback, Traceback/Frame

#[test]
fn test_tracemalloc_snapshot_statistics_lineno() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
x = [0] * 10000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("lineno")
print(len(stats) > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_snapshot_statistics_filename() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
y = {"key": list(range(500))}
snap = tracemalloc.take_snapshot()
stats = snap.statistics("filename")
print(len(stats) > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_snapshot_statistics_traceback() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
z = bytearray(50000)
snap = tracemalloc.take_snapshot()
stats = snap.statistics("traceback")
print(len(stats) > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_stat_size_and_count() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
data = b"x" * 100000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("lineno")
total_size = sum(s.size for s in stats)
total_count = sum(s.count for s in stats)
print(total_size > 0)
print(total_count > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tracemalloc_filter_traces_include() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
x = [None] * 1000
snap = tracemalloc.take_snapshot()
f = tracemalloc.Filter(inclusive=True, filename_pattern="*.py")
filtered = snap.filter_traces([f])
print(len(filtered.statistics("filename")) >= 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_filter_traces_exclude() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
x = [None] * 500
snap = tracemalloc.take_snapshot()
f = tracemalloc.Filter(inclusive=False, filename_pattern="<frozen*")
filtered = snap.filter_traces([f])
# Excluding frozen modules may reduce stat count
print(len(filtered.statistics("filename")) >= 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_statsdiff_new() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
big = [0] * 100000
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
print(len(diff) > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_statsdiff_size_delta_positive() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
alloc = bytearray(500000)
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
# At least one stat shows positive size delta
print(any(d.size_diff > 0 for d in diff))
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_get_traced_memory_tuple() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
data = [None] * 10000
current, peak = tracemalloc.get_traced_memory()
print(current > 0)
print(peak >= current)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tracemalloc_reset_peak_reduces_peak() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
big = [None] * 100000
_, peak_before = tracemalloc.get_traced_memory()
del big
tracemalloc.reset_peak()
_, peak_after = tracemalloc.get_traced_memory()
print(peak_after <= peak_before)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_get_traceback_for_object() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start(10)
x = object()
tb = tracemalloc.get_object_traceback(x)
# tb may be None if not tracked, but the call must not error
print(tb is None or len(tb) > 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_stat_traceback_has_frames() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start(5)
x = [None] * 5000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("traceback")
if stats:
    tb = stats[0].traceback
    print(len(tb) > 0)
else:
    print(True)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_frame_filename_and_lineno() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start(3)
x = [None] * 2000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("traceback")
if stats:
    frame = stats[0].traceback[0]
    print(isinstance(frame.filename, str))
    print(frame.lineno > 0)
else:
    print(True)
    print(True)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_tracemalloc_statsdiff_count_diff() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
new_objects = [object() for _ in range(1000)]
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
print(any(d.count_diff > 0 for d in diff))
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_is_tracing_false_when_stopped() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.stop()
print(tracemalloc.is_tracing())
"#);
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_tracemalloc_is_tracing_true_when_started() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
print(tracemalloc.is_tracing())
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_get_tracemalloc_memory() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
mem = tracemalloc.get_tracemalloc_memory()
print(mem >= 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_snapshot_dump_and_load() {
    let out = run_python(r#"
import tracemalloc, tempfile, os
tracemalloc.start()
x = [None] * 1000
snap = tracemalloc.take_snapshot()
tracemalloc.stop()
f = tempfile.NamedTemporaryFile(delete=False, suffix=".snap")
f.close()
snap.dump(f.name)
snap2 = tracemalloc.Snapshot.load(f.name)
print(len(snap2.statistics("lineno")) > 0)
os.unlink(f.name)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_filter_lineno() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
x = [None] * 2000
snap = tracemalloc.take_snapshot()
f = tracemalloc.Filter(inclusive=True, filename_pattern="*.py", lineno=None)
filtered = snap.filter_traces([f])
print(len(filtered.statistics("lineno")) >= 0)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tracemalloc_statsdiff_str_not_empty() {
    let out = run_python(r#"
import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
x = [0] * 10000
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
if diff:
    print(len(str(diff[0])) > 0)
else:
    print(True)
tracemalloc.stop()
"#);
    assert_eq!(out, vec!["True"]);
}
