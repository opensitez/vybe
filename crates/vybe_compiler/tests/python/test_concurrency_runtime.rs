//! threading, queue, subprocess, multiprocessing basics.

crate::runtime_case!(
    threading_thread_class,
    "import threading\nprint(callable(threading.Thread))\n",
    "True"
);
crate::runtime_case!(
    threading_lock,
    "import threading\nlock = threading.Lock()\nprint(lock.acquire())\n",
    "True"
);
crate::runtime_case!(
    threading_rlock,
    "import threading\nlock = threading.RLock()\nprint(lock.acquire())\n",
    "True"
);
crate::runtime_case!(
    threading_event,
    "import threading\ne = threading.Event()\nprint(e.is_set())\n",
    "False"
);
crate::runtime_case!(
    threading_event_set,
    "import threading\ne = threading.Event()\ne.set()\nprint(e.is_set())\n",
    "True"
);
crate::runtime_case!(
    threading_condition,
    "import threading\nc = threading.Condition()\nprint(c.acquire())\n",
    "True"
);
crate::runtime_case!(
    threading_semaphore,
    "import threading\ns = threading.Semaphore(2)\nprint(s.acquire())\n",
    "True"
);
crate::runtime_case!(
    threading_bounded_semaphore,
    "import threading\ns = threading.BoundedSemaphore(1)\nprint(s.acquire())\n",
    "True"
);
crate::runtime_case!(
    threading_barrier,
    "import threading\nb = threading.Barrier(1)\nprint(b.parties)\n",
    "1"
);
crate::runtime_case!(
    threading_local,
    "import threading\nl = threading.local()\nl.x = 1\nprint(l.x)\n",
    "1"
);
crate::runtime_case!(
    threading_active_count,
    "import threading\nprint(threading.active_count() >= 1)\n",
    "True"
);
crate::runtime_case!(
    threading_current_thread,
    "import threading\nprint(isinstance(threading.current_thread(), threading.Thread))\n",
    "True"
);
crate::runtime_case!(
    threading_main_thread,
    "import threading\nprint(threading.current_thread() is threading.main_thread())\n",
    "True"
);
crate::runtime_case!(
    threading_get_ident,
    "import threading\nprint(threading.get_ident() > 0)\n",
    "True"
);
crate::runtime_case!(
    queue_queue,
    "import queue\nq = queue.Queue()\nq.put(1)\nprint(q.get())\n",
    "1"
);
crate::runtime_case!(
    queue_lifo,
    "import queue\nq = queue.LifoQueue()\nq.put(1)\nq.put(2)\nprint(q.get())\n",
    "2"
);
crate::runtime_case!(
    queue_priority,
    "import queue\nq = queue.PriorityQueue()\nq.put((1, 'a'))\nq.put((2, 'b'))\nprint(q.get()[1])\n",
    "a"
);
crate::runtime_case!(
    queue_empty,
    "import queue\nq = queue.Queue()\nprint(q.empty())\n",
    "True"
);
crate::runtime_case!(
    queue_full,
    "import queue\nq = queue.Queue(maxsize=1)\nq.put(1)\nprint(q.full())\n",
    "True"
);
crate::runtime_case!(
    queue_qsize,
    "import queue\nq = queue.Queue()\nq.put(1)\nprint(q.qsize())\n",
    "1"
);
crate::runtime_case!(
    subprocess_popen_class,
    "import subprocess\nprint(callable(subprocess.Popen))\n",
    "True"
);
crate::runtime_case!(
    subprocess_run_exists,
    "import subprocess\nprint(callable(subprocess.run))\n",
    "True"
);
crate::runtime_case!(
    subprocess_call_exists,
    "import subprocess\nprint(callable(subprocess.call))\n",
    "True"
);
crate::runtime_case!(
    subprocess_check_output_exists,
    "import subprocess\nprint(callable(subprocess.check_output))\n",
    "True"
);
crate::runtime_case!(
    subprocess_pipe,
    "import subprocess\nprint(subprocess.PIPE is not None)\n",
    "True"
);
crate::runtime_case!(
    subprocess_devnull,
    "import subprocess\nprint(subprocess.DEVNULL is not None)\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_process,
    "import multiprocessing\nprint(callable(multiprocessing.Process))\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_queue,
    "import multiprocessing\nprint(callable(multiprocessing.Queue))\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_pipe,
    "import multiprocessing\nprint(callable(multiprocessing.Pipe))\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_value,
    "import multiprocessing\nv = multiprocessing.Value('i', 0)\nprint(v.value)\n",
    "0"
);
crate::runtime_case!(
    multiprocessing_array,
    "import multiprocessing\na = multiprocessing.Array('i', [1, 2, 3])\nprint(a[1])\n",
    "2"
);
crate::runtime_case!(
    multiprocessing_lock,
    "import multiprocessing\nprint(callable(multiprocessing.Lock))\n",
    "True"
);
crate::runtime_case!(
    concurrent_futures_thread,
    "import concurrent.futures\nprint(hasattr(concurrent.futures, 'ThreadPoolExecutor'))\n",
    "True"
);
crate::runtime_case!(
    concurrent_futures_process,
    "import concurrent.futures\nprint(hasattr(concurrent.futures, 'ProcessPoolExecutor'))\n",
    "True"
);
crate::runtime_case!(
    threading_timer,
    "import threading\nprint(callable(threading.Timer))\n",
    "True"
);
crate::runtime_case!(
    threading_excepthook,
    "import threading\nprint(hasattr(threading, 'excepthook'))\n",
    "True"
);
crate::runtime_case!(
    queue_simplequeue,
    "import queue\nprint(hasattr(queue, 'SimpleQueue'))\n",
    "True"
);
crate::runtime_case!(
    subprocess_completedprocess,
    "import subprocess\nprint(hasattr(subprocess, 'CompletedProcess'))\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_cpu_count,
    "import multiprocessing\nprint(multiprocessing.cpu_count() >= 1)\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_current_process,
    "import multiprocessing\nprint(multiprocessing.current_process().name)\n",
    "MainProcess"
);
crate::runtime_case!(
    threading_enumerate,
    "import threading\nprint(len(threading.enumerate()) >= 1)\n",
    "True"
);
crate::runtime_case!(
    queue_joinable,
    "import queue\nprint(hasattr(queue.Queue, 'join'))\n",
    "True"
);
crate::runtime_case!(
    subprocess_timeout,
    "import subprocess\nprint(hasattr(subprocess, 'TimeoutExpired'))\n",
    "True"
);
crate::runtime_case!(
    multiprocessing_manager,
    "import multiprocessing\nprint(callable(multiprocessing.Manager))\n",
    "True"
);
crate::runtime_case!(
    concurrent_futures_wait,
    "import concurrent.futures\nprint(callable(concurrent.futures.wait))\n",
    "True"
);
crate::runtime_case!(
    concurrent_futures_as_completed,
    "import concurrent.futures\nprint(callable(concurrent.futures.as_completed))\n",
    "True"
);
crate::runtime_case!(
    threading_stack_size,
    "import threading\nprint(threading.stack_size() > 0)\n",
    "True"
);
crate::runtime_case!(
    queue_task_done,
    "import queue\nq = queue.Queue()\nq.put(1)\nq.get()\nq.task_done()\nprint('ok')\n",
    "ok"
);

crate::compile_case!(threading_thread_start, "import threading\nt = threading.Thread(target=lambda: None)\nt.start()\nt.join()\n");
crate::compile_case!(subprocess_run_echo, "import subprocess\nsubprocess.run(['echo', 'hi'], capture_output=True)\n");
crate::compile_case!(multiprocessing_pool, "import multiprocessing\nmultiprocessing.Pool\n");
crate::compile_case!(concurrent_futures_submit, "import concurrent.futures\nwith concurrent.futures.ThreadPoolExecutor() as ex:\n ex.submit(lambda: 1)\n");
crate::compile_case!(multiprocessing_spawn, "import multiprocessing\nmultiprocessing.get_context('spawn')\n");
