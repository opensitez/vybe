use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Multiprocessing Pool & Processes — Process, Pool.map, Queue, Pipe, Value, Array, Manager
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_multiprocessing_process_creation_join() {
    let src = r#"
import multiprocessing

def worker(q):
    q.put("hello from worker")

if __name__ == "__main__":
    q = multiprocessing.Queue()
    p = multiprocessing.Process(target=worker, args=(q,))
    p.start()
    p.join()
    print(q.get())
"#;
    assert_eq!(run_python(src), vec!["hello from worker"]);
}

#[test]
fn test_py_multiprocessing_pool_map_parallel() {
    let src = r#"
import multiprocessing

def square(x):
    return x * x

if __name__ == "__main__":
    with multiprocessing.Pool(processes=2) as pool:
        results = pool.map(square, [1, 2, 3, 4])
    print(results)
"#;
    assert_eq!(run_python(src), vec!["[1, 4, 9, 16]"]);
}

#[test]
fn test_py_multiprocessing_pipe_ipc() {
    let src = r#"
import multiprocessing

def child_proc(conn):
    msg = conn.recv()
    conn.send(f"ACK: {msg}")
    conn.close()

if __name__ == "__main__":
    parent_conn, child_conn = multiprocessing.Pipe()
    p = multiprocessing.Process(target=child_proc, args=(child_conn,))
    p.start()
    parent_conn.send("ping")
    reply = parent_conn.recv()
    p.join()
    print(reply)
"#;
    assert_eq!(run_python(src), vec!["ACK: ping"]);
}

#[test]
fn test_py_multiprocessing_shared_value_and_array() {
    let src = r#"
import multiprocessing

def f(n, a):
    n.value = 3.14159
    for i in range(len(a)):
        a[i] = -a[i]

if __name__ == "__main__":
    num = multiprocessing.Value('d', 0.0)
    arr = multiprocessing.Array('i', range(5))

    p = multiprocessing.Process(target=f, args=(num, arr))
    p.start()
    p.join()

    print(round(num.value, 2))
    print(list(arr))
"#;
    assert_eq!(run_python(src), vec!["3.14", "[0, -1, -2, -3, -4]"]);
}

#[test]
fn test_py_multiprocessing_manager_dict_list() {
    let src = r#"
import multiprocessing

def worker(d, l):
    d["status"] = "ok"
    l.append(42)

if __name__ == "__main__":
    with multiprocessing.Manager() as manager:
        d = manager.dict()
        l = manager.list()

        p = multiprocessing.Process(target=worker, args=(d, l))
        p.start()
        p.join()

        print(d["status"])
        print(list(l))
"#;
    assert_eq!(run_python(src), vec!["ok", "[42]"]);
}

#[test]
fn test_py_multiprocessing_pool_starmap() {
    let src = r#"
import multiprocessing

def add(a, b):
    return a + b

if __name__ == "__main__":
    with multiprocessing.Pool(2) as pool:
        res = pool.starmap(add, [(1, 2), (3, 4), (5, 6)])
    print(res)
"#;
    assert_eq!(run_python(src), vec!["[3, 7, 11]"]);
}

#[test]
fn test_py_multiprocessing_cpu_count() {
    let src = r#"
import multiprocessing

cpus = multiprocessing.cpu_count()
print(isinstance(cpus, int) and cpus > 0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_multiprocessing_pool_apply_async() {
    let src = r#"
import multiprocessing

def multiply(x, y):
    return x * y

if __name__ == "__main__":
    with multiprocessing.Pool(2) as pool:
        res1 = pool.apply_async(multiply, (3, 4))
        res2 = pool.apply_async(multiply, (5, 6))
        print(res1.get())
        print(res2.get())
"#;
    assert_eq!(run_python(src), vec!["12", "30"]);
}

#[test]
fn test_py_multiprocessing_current_process_name() {
    let src = r#"
import multiprocessing

def worker(q):
    q.put(multiprocessing.current_process().name)

if __name__ == "__main__":
    q = multiprocessing.Queue()
    p = multiprocessing.Process(target=worker, args=(q,), name="CustomWorker")
    p.start()
    p.join()
    print(q.get())
"#;
    assert_eq!(run_python(src), vec!["CustomWorker"]);
}

#[test]
fn test_py_multiprocessing_active_children() {
    let src = r#"
import multiprocessing

if __name__ == "__main__":
    children = multiprocessing.active_children()
    print(isinstance(children, list))
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
