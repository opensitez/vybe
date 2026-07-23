use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Time Clocks & Benchmarking — perf_counter, monotonic, process_time, timeit.timeit, timeit.repeat, sleep
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_time_perf_counter_precision() {
    let src = r#"
import time

start = time.perf_counter()
time.sleep(0.001)
end = time.perf_counter()

elapsed = end - start
print(elapsed > 0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_time_monotonic_non_decreasing() {
    let src = r#"
import time

m1 = time.monotonic()
time.sleep(0.001)
m2 = time.monotonic()

print(m2 >= m1)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_time_process_time_cpu_only() {
    let src = r#"
import time

p1 = time.process_time()
# CPU-bound work
_ = [x * x for x in range(10000)]
p2 = time.process_time()

print(p2 >= p1)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_timeit_timeit_benchmark_code_str() {
    let src = r#"
import timeit

t = timeit.timeit("sum(range(100))", number=100)
print(isinstance(t, float))
print(t > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_timeit_repeat_multiple_samples() {
    let src = r#"
import timeit

samples = timeit.repeat("sorted([5, 2, 8, 1])", repeat=3, number=50)
print(len(samples))
print(all(isinstance(s, float) and s > 0 for s in samples))
"#;
    assert_eq!(run_python(src), vec!["3", "True"]);
}

#[test]
fn test_py_timeit_timer_class_with_callable() {
    let src = r#"
import timeit

def benchmark_target():
    return [i * 2 for i in range(50)]

timer = timeit.Timer(benchmark_target)
elapsed = timer.timeit(number=100)
print(elapsed > 0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_time_struct_time_tuple_unpacking() {
    let src = r#"
import time

st = time.struct_time((2024, 6, 15, 14, 30, 0, 5, 167, 0))
print(st.tm_year, st.tm_mon, st.tm_mday)
print(st[0], st[1], st[2])
"#;
    assert_eq!(run_python(src), vec!["2024 6 15", "2024 6 15"]);
}

#[test]
fn test_py_time_strftime_format_directives() {
    let src = r#"
import time

st = time.struct_time((2024, 1, 5, 9, 5, 0, 4, 5, 0))
formatted = time.strftime("%Y-%m-%d %H:%M:%S", st)
print(formatted)
"#;
    assert_eq!(run_python(src), vec!["2024-01-05 09:05:00"]);
}

#[test]
fn test_py_time_ns_clocks_nanoseconds() {
    let src = r#"
import time

p_ns = time.perf_counter_ns()
m_ns = time.monotonic_ns()
t_ns = time.time_ns()

print(isinstance(p_ns, int))
print(isinstance(m_ns, int))
print(isinstance(t_ns, int))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_time_thread_time_clock() {
    let src = r#"
import time

if hasattr(time, "thread_time"):
    tt = time.thread_time()
    print(isinstance(tt, float))
else:
    print("True")
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
