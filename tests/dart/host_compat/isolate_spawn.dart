// vybe-test: dart/host_compat/isolate_spawn
// origin: languages/dart/tests/dart/test_host_compat.rs

void worker(msg) {} Isolate.spawn(worker, 'hello');

void main() {}
