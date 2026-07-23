use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: time + timeit + profile — time measuring, benchmarking, sleep, struct_time
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_time_perf_counter_monotonic() {
    let src = r#"
import time

t1 = time.perf_counter()
t2 = time.monotonic()
time.sleep(0.01)
dt1 = time.perf_counter() - t1
dt2 = time.monotonic() - t2

print(dt1 > 0)
print(dt2 > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_time_struct_time_and_gmtime() {
    let src = r#"
import time

t = time.gmtime(1700000000)
print(t.tm_year)
print(t.tm_mon)
print(t.tm_mday)
print(isinstance(t, time.struct_time))
"#;
    assert_eq!(run_python(src), vec!["2023", "11", "14", "True"]);
}

#[test]
fn test_py_time_strftime_strptime() {
    let src = r#"
import time

t_struct = time.struct_time((2024, 6, 15, 12, 30, 0, 5, 167, 0))
formatted = time.strftime("%Y-%m-%d %H:%M:%S", t_struct)
print(formatted)

parsed = time.strptime("2024-06-15 12:30:00", "%Y-%m-%d %H:%M:%S")
print(parsed.tm_year, parsed.tm_mon, parsed.tm_mday)
"#;
    assert_eq!(run_python(src), vec!["2024-06-15 12:30:00", "2024 6 15"]);
}

#[test]
fn test_py_timeit_basic_benchmark() {
    let src = r#"
import timeit

t1 = timeit.timeit("sum(range(100))", number=1000)
t2 = timeit.timeit("[x for x in range(100)]", number=1000)
print(t1 > 0)
print(t2 > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_timeit_repeat() {
    let src = r#"
import timeit

runs = timeit.repeat("sorted([3, 1, 4, 1, 5])", number=100, repeat=3)
print(len(runs))
print(all(r > 0 for r in runs))
"#;
    assert_eq!(run_python(src), vec!["3", "True"]);
}

#[test]
fn test_py_time_process_time() {
    let src = r#"
import time

t0 = time.process_time()
_ = sum(i * i for i in range(10000))
t1 = time.process_time()
print(t1 >= t0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_time_timezone_and_altzone() {
    let src = r#"
import time

print(isinstance(time.timezone, int))
print(isinstance(time.tzname, tuple))
print(len(time.tzname) == 2)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_time_ctime_mktime() {
    let src = r#"
import time

t_struct = time.gmtime(1700000000)
ts = time.mktime(t_struct)
print(isinstance(ts, float))

s = time.ctime(1700000000)
print(isinstance(s, str))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_cprofile_runctx() {
    let src = r#"
import cProfile, pstats, io

def fib(n):
    return n if n < 2 else fib(n-1) + fib(n-2)

pr = cProfile.Profile()
pr.enable()
fib(10)
pr.disable()

s = io.StringIO()
ps = pstats.Stats(pr, stream=s).sort_stats('cumulative')
ps.print_stats(5)

output = s.getvalue()
print("function calls" in output)
print("fib" in output)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_time_ns_nanosecond_clocks() {
    let src = r#"
import time

t_ns = time.perf_counter_ns()
mono_ns = time.monotonic_ns()
time_ns = time.time_ns()

print(isinstance(t_ns, int))
print(isinstance(mono_ns, int))
print(isinstance(time_ns, int))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}
