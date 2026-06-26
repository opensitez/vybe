//! Future, async functions, and await — core async semantics.

dart_cases! {
    future_value_int_is_non_null => {
        r#"void main() {
  var f = Future.value(42);
  print(f != null);
}"#,
        ["true"]
    };

    future_value_string_stored => {
        r#"void main() {
  var f = Future.value('ok');
  print(f != null);
}"#,
        ["true"]
    };

    future_value_bool_literal => {
        r#"void main() {
  var f = Future.value(true);
  print(f != null);
}"#,
        ["true"]
    };

    future_sync_completes_with_value => {
        r#"void main() {
  var f = Future.sync(() => 7);
  print(f != null);
}"#,
        ["true"]
    };

    future_sync_returns_string => {
        r#"void main() {
  var f = Future.sync(() => 'sync');
  print(f != null);
}"#,
        ["true"]
    };

    async_function_returns_future_int => {
        r#"Future<int> loadCount() async {
  return 10;
}
void main() {
  var f = loadCount();
  print(f != null);
}"#,
        ["true"]
    };

    async_function_returns_future_string => {
        r#"Future<String> loadName() async {
  return 'Ada';
}
void main() {
  var f = loadName();
  print(f != null);
}"#,
        ["true"]
    };

    async_void_function_returns_future => {
        r#"Future<void> noop() async {}
void main() {
  var f = noop();
  print(f != null);
}"#,
        ["true"]
    };

    async_method_on_class => {
        r#"class Api {
  Future<int> fetch() async {
    return 3;
  }
}
void main() {
  var f = Api().fetch();
  print(f != null);
}"#,
        ["true"]
    };

    async_arrow_function => {
        r#"void main() {
  Future<int> f() async => 5;
  print(f() != null);
}"#,
        ["true"]
    };

    await_simple_int_result => {
        r#"Future<int> twice(int n) async {
  return n * 2;
}
void main() async {
  var v = await twice(11);
  print(v);
}"#,
        ["22"]
    };

    await_simple_string_result => {
        r#"Future<String> label() async {
  return 'ready';
}
void main() async {
  var s = await label();
  print(s);
}"#,
        ["ready"]
    };

    await_local_async_function => {
        r#"void main() async {
  Future<int> bump() async {
    return 4;
  }
  print(await bump());
}"#,
        ["4"]
    };

    await_chained_async_calls => {
        r#"Future<int> step1() async {
  return 2;
}
Future<int> step2(int n) async {
  return n + 3;
}
void main() async {
  var a = await step1();
  var b = await step2(a);
  print(b);
}"#,
        ["5"]
    };

    await_in_if_branch => {
        r#"Future<int> pick(bool useBig) async {
  return useBig ? 100 : 1;
}
void main() async {
  if (true) {
    print(await pick(true));
  } else {
    print(await pick(false));
  }
}"#,
        ["100"]
    };

    await_in_for_loop_accumulates => {
        r#"Future<int> next(int n) async {
  return n + 1;
}
void main() async {
  var sum = 0;
  for (var i = 0; i < 3; i++) {
    sum = sum + await next(i);
  }
  print(sum);
}"#,
        ["6"]
    };

    await_expression_in_addition => {
        r#"Future<int> five() async {
  return 5;
}
void main() async {
  print(2 + await five());
}"#,
        ["7"]
    };

    await_void_then_prints_after => {
        r#"Future<void> markDone() async {}
void main() async {
  await markDone();
  print('done');
}"#,
        ["done"]
    };

    async_class_method_awaited => {
        r#"class Worker {
  Future<int> produce() async {
    return 9;
  }
}
void main() async {
  var w = Worker();
  print(await w.produce());
}"#,
        ["9"]
    };

    nested_async_function_await => {
        r#"Future<int> outer() async {
  Future<int> inner() async {
    return 6;
  }
  return await inner();
}
void main() async {
  print(await outer());
}"#,
        ["6"]
    };

    future_then_prints_value => {
        r#"void main() {
  Future.value(15).then((v) {
    print(v);
  });
}"#,
        ["15"]
    };

    future_then_chain_doubles => {
        r#"void main() {
  Future.value(3)
      .then((v) => v * 2)
      .then((v) {
    print(v);
  });
}"#,
        ["6"]
    };

    future_catch_error_prints_message => {
        r#"void main() {
  Future<int>.error('fail').catchError((e) {
    print(e);
  });
}"#,
        ["fail"]
    };

    future_when_complete_runs_callback => {
        r#"void main() {
  Future.value(1).whenComplete(() {
    print('finished');
  });
}"#,
        ["finished"]
    };

    async_return_after_local_var => {
        r#"Future<int> compute() async {
  var base = 8;
  return base + 1;
}
void main() async {
  print(await compute());
}"#,
        ["9"]
    };

    await_bool_future => {
        r#"Future<bool> flag() async {
  return false;
}
void main() async {
  print(await flag());
}"#,
        ["false"]
    };

    async_with_early_return => {
        r#"Future<int> early(bool skip) async {
  if (skip) {
    return 0;
  }
  return 99;
}
void main() async {
  print(await early(false));
}"#,
        ["99"]
    };

    await_twice_same_future_fn => {
        r#"Future<int> constant() async {
  return 12;
}
void main() async {
  var a = await constant();
  var b = await constant();
  print(a + b);
}"#,
        ["24"]
    };

    async_parameter_passed_through => {
        r#"Future<int> echo(int n) async {
  return n;
}
void main() async {
  print(await echo(17));
}"#,
        ["17"]
    };

    future_value_zero => {
        r#"void main() {
  var f = Future.value(0);
  print(f != null);
}"#,
        ["true"]
    };

    async_while_loop_with_await => {
        r#"Future<int> tick(int n) async {
  return n;
}
void main() async {
  var i = 0;
  var total = 0;
  while (i < 2) {
    total = total + await tick(i + 1);
    i = i + 1;
  }
  print(total);
}"#,
        ["3"]
    };

    await_nullable_int_coalesced => {
        r#"Future<int?> maybe() async {
  return null;
}
void main() async {
  var v = await maybe();
  print(v ?? 5);
}"#,
        ["5"]
    };

    async_try_finally_still_returns => {
        r#"Future<int> guarded() async {
  try {
    return 2;
  } finally {
    print('cleanup');
  }
}
void main() async {
  print(await guarded());
}"#,
        ["cleanup", "2"]
    };

    future_microtask_created => {
        r#"void main() {
  var f = Future.microtask(() => 1);
  print(f != null);
}"#,
        ["true"]
    };

    async_multiple_returns_in_branches => {
        r#"Future<String> branch(int n) async {
  if (n > 0) {
    return 'pos';
  }
  return 'neg';
}
void main() async {
  print(await branch(1));
}"#,
        ["pos"]
    };

    await_in_ternary_expression => {
        r#"Future<int> left() async {
  return 1;
}
Future<int> right() async {
  return 2;
}
void main() async {
  var pickLeft = true;
  print(await (pickLeft ? left() : right()));
}"#,
        ["1"]
    };

    async_static_method_style => {
        r#"class MathBox {
  static Future<int> triple(int n) async {
    return n * 3;
  }
}
void main() async {
  print(await MathBox.triple(4));
}"#,
        ["12"]
    };

    future_delayed_reference => {
        r#"void main() {
  var f = Future.delayed(Duration(milliseconds: 1), () => 1);
  print(f != null);
}"#,
        ["true"]
    };

    async_recursive_base_case => {
        r#"Future<int> down(int n) async {
  if (n <= 0) {
    return 0;
  }
  return n + await down(n - 1);
}
void main() async {
  print(await down(3));
}"#,
        ["6"]
    };

    await_string_interpolation_result => {
        r#"Future<String> name() async {
  return 'dart';
}
void main() async {
  var n = await name();
  print('lang=$n');
}"#,
        ["lang=dart"]
    };
}
