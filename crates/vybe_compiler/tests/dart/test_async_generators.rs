use super::helpers::compile_ok;

// ── async functions returning Future ─────────────────────────

#[test]
fn async_fn_int() {
    compile_ok("Future<int> fetchCount() async { return 42; }");
}
#[test]
fn async_fn_string() {
    compile_ok("Future<String> fetchName() async { return 'Alice'; }");
}
#[test]
fn async_fn_void() {
    compile_ok("Future<void> doWork() async { print('working'); }");
}
#[test]
fn async_method() {
    compile_ok("class Api { Future<String> get(String url) async { return '{}'; } }");
}

#[test]
fn async_await_call() {
    compile_ok(
        r#"
Future<int> compute() async { return 99; }
Future<void> main() async { var result = await compute(); print(result); }
"#,
    );
}

#[test]
fn async_await_in_loop() {
    compile_ok(
        r#"
Future<int> fetch(int i) async { return i * 2; }
Future<void> main() async {
  for (var i = 0; i < 3; i++) {
    var v = await fetch(i);
    print(v);
  }
}
"#,
    );
}

#[test]
fn async_try_catch() {
    compile_ok(
        r#"
Future<void> risky() async { throw Exception('async error'); }
Future<void> main() async {
  try { await risky(); } catch (e) { print('caught'); }
}
"#,
    );
}

// ── Future combinators ───────────────────────────────────────

#[test]
fn future_value() {
    compile_ok("var f = Future.value(42);");
}
#[test]
fn future_error() {
    compile_ok("var f = Future<int>.error(Exception('fail'));");
}
#[test]
fn future_delayed() {
    compile_ok("var f = Future.delayed(Duration(milliseconds: 100), () => 42);");
}
#[test]
fn future_wait() {
    compile_ok(
        "var f1 = Future.value(1); var f2 = Future.value(2); var all = Future.wait([f1, f2]);",
    );
}
#[test]
fn future_any() {
    compile_ok(
        "var f1 = Future.value(1); var f2 = Future.value(2); var first = Future.any([f1, f2]);",
    );
}
#[test]
fn future_microtask() {
    compile_ok("var f = Future.microtask(() => 42);");
}
#[test]
fn future_sync() {
    compile_ok("var f = Future.sync(() => 42);");
}

#[test]
fn future_then() {
    compile_ok("Future.value(1).then((v) => print(v));");
}
#[test]
fn future_catch_error() {
    compile_ok("Future.error('oops').catchError((e) => print(e));");
}
#[test]
fn future_when_complete() {
    compile_ok("Future.value(1).whenComplete(() => print('done'));");
}
#[test]
fn future_then_chain() {
    compile_ok("Future.value(1).then((v) => v * 2).then((v) => print(v));");
}

// ── async* generators ────────────────────────────────────────

#[test]
fn async_star_basic() {
    compile_ok("Stream<int> count() async* { yield 1; yield 2; yield 3; }");
}

#[test]
fn async_star_loop() {
    compile_ok(
        r#"
Stream<int> range(int n) async* {
  for (var i = 0; i < n; i++) { yield i; }
}
"#,
    );
}

#[test]
fn async_star_yield_star() {
    compile_ok(
        r#"
Stream<int> first() async* { yield 1; yield 2; }
Stream<int> combined() async* { yield* first(); yield 3; }
"#,
    );
}

#[test]
fn async_star_with_transform() {
    compile_ok(
        r#"
Stream<String> messages() async* {
  var items = ['a', 'b', 'c'];
  for (var item in items) { yield item.toUpperCase(); }
}
"#,
    );
}

// ── sync* generators ─────────────────────────────────────────

#[test]
fn sync_star_basic() {
    compile_ok("Iterable<int> count() sync* { yield 1; yield 2; yield 3; }");
}

#[test]
fn sync_star_loop() {
    compile_ok(
        r#"
Iterable<int> range(int end) sync* {
  for (var i = 0; i < end; i++) { yield i; }
}
"#,
    );
}

#[test]
fn sync_star_yield_star() {
    compile_ok(
        r#"
Iterable<int> first() sync* { yield 1; yield 2; }
Iterable<int> all() sync* { yield* first(); yield 3; }
"#,
    );
}

#[test]
fn sync_star_fibonacci() {
    compile_ok(
        r#"
Iterable<int> fibs() sync* {
  int a = 0, b = 1;
  while (true) { yield a; var c = a + b; a = b; b = c; }
}
"#,
    );
}

#[test]
fn sync_star_used() {
    compile_ok(
        "Iterable<int> evens(int n) sync* { for (var i = 0; i < n; i += 2) yield i; } void main() { var list = evens(10).toList(); }",
    );
}

// ── yield (standalone) ───────────────────────────────────────

#[test]
fn yield_in_sync_star() {
    compile_ok("Iterable<int> gen() sync* { yield 1; yield 2; yield 3; }");
}

#[test]
fn yield_in_async_star() {
    compile_ok("Stream<int> gen() async* { yield 1; yield 2; yield 3; }");
}

#[test]
fn yield_star() {
    compile_ok(
        "Iterable<int> a() sync* { yield 1; } Iterable<int> b() sync* { yield* a(); yield 2; }",
    );
}

// ── await for ────────────────────────────────────────────────

#[test]
fn await_for_stream() {
    compile_ok(
        r#"
Stream<int> nums() async* { yield 1; yield 2; yield 3; }
Future<void> main() async {
  await for (var n in nums()) {
    print(n);
  }
}
"#,
    );
}

#[test]
fn await_for_break() {
    compile_ok(
        r#"
Stream<int> nums() async* { for (var i = 0; i < 10; i++) yield i; }
Future<void> main() async {
  await for (var n in nums()) {
    if (n > 3) break;
    print(n);
  }
}
"#,
    );
}

// ── Stream construction ───────────────────────────────────────

#[test]
fn stream_from_iterable() {
    compile_ok("var s = Stream.fromIterable([1, 2, 3]);");
}
#[test]
fn stream_value() {
    compile_ok("var s = Stream.value(42);");
}
#[test]
fn stream_empty() {
    compile_ok("var s = Stream<int>.empty();");
}
#[test]
fn stream_periodic() {
    compile_ok("var s = Stream.periodic(Duration(seconds: 1), (i) => i);");
}

#[test]
fn stream_map_method() {
    compile_ok("var s = Stream.fromIterable([1, 2, 3]).map((x) => x * 2);");
}
#[test]
fn stream_where_method() {
    compile_ok("var s = Stream.fromIterable([1, 2, 3, 4]).where((x) => x > 2);");
}
#[test]
fn stream_to_list() {
    compile_ok("Future<List<int>> f() async => Stream.fromIterable([1, 2, 3]).toList();");
}

// ── Duration ─────────────────────────────────────────────────

#[test]
fn duration_millis() {
    compile_ok("var d = Duration(milliseconds: 500);");
}
#[test]
fn duration_seconds() {
    compile_ok("var d = Duration(seconds: 30);");
}
#[test]
fn duration_minutes() {
    compile_ok("var d = Duration(minutes: 5);");
}
#[test]
fn duration_hours() {
    compile_ok("var d = Duration(hours: 2);");
}
#[test]
fn duration_zero() {
    compile_ok("var d = Duration.zero;");
}
