//! Closures: capturing outer variables, mutation, and passing closures as arguments.

dart_cases! {
    closure_reads_captured_local_integer => {
        r#"void main() {
  var x = 10;
  var fn = () => x;
  print(fn());
}"#,
        ["10"]
    };

    closure_computes_with_captured_multiplier => {
        r#"void main() {
  var factor = 3;
  var scale = (int n) => n * factor;
  print(scale(7));
}"#,
        ["21"]
    };

    closure_captures_string_and_concatenates => {
        r#"void main() {
  var prefix = 'vybe';
  var label = (String suffix) => '$prefix-$suffix';
  print(label('dart'));
}"#,
        ["vybe-dart"]
    };

    closure_captures_bool_flag => {
        r#"void main() {
  var enabled = true;
  var check = () => enabled;
  print(check());
}"#,
        ["true"]
    };

    closure_mutates_captured_counter => {
        r#"void main() {
  var count = 0;
  var inc = () {
    count++;
  };
  inc();
  inc();
  inc();
  print(count);
}"#,
        ["3"]
    };

    closure_decrements_captured_value => {
        r#"void main() {
  var remaining = 5;
  var step = () {
    remaining--;
  };
  step();
  step();
  print(remaining);
}"#,
        ["3"]
    };

    closure_appends_to_captured_list => {
        r#"void main() {
  var items = <int>[];
  var push = (int v) {
    items.add(v);
  };
  push(1);
  push(2);
  print(items.join(','));
}"#,
        ["1,2"]
    };

    closure_as_argument_to_apply_helper => {
        r#"int apply(int x, int Function(int) fn) {
  return fn(x);
}
void main() {
  print(apply(5, (n) => n + 1));
}"#,
        ["6"]
    };

    closure_as_argument_doubles_input => {
        r#"int apply(int x, int Function(int) fn) {
  return fn(x);
}
void main() {
  print(apply(9, (n) => n * 2));
}"#,
        ["18"]
    };

    closure_as_argument_to_void_consumer => {
        r#"void consume(int x, void Function(int) fn) {
  fn(x);
}
void main() {
  var seen = 0;
  consume(7, (n) {
    seen = n;
  });
  print(seen);
}"#,
        ["7"]
    };

    make_adder_closure_captures_offset => {
        r#"int Function(int) makeAdder(int n) {
  return (x) => x + n;
}
void main() {
  var add5 = makeAdder(5);
  print(add5(10));
}"#,
        ["15"]
    };

    make_multiplier_closure_captures_factor => {
        r#"int Function(int) makeMultiplier(int m) {
  return (x) => x * m;
}
void main() {
  var triple = makeMultiplier(3);
  print(triple(4));
}"#,
        ["12"]
    };

    closure_captures_outer_function_parameter => {
        r#"int wrap(int base) {
  var add = (int n) => base + n;
  return add(4);
}
void main() {
  print(wrap(10));
}"#,
        ["14"]
    };

    closure_returns_another_closure => {
        r#"Function makeChain(int a) {
  return (int b) {
    return (int c) => a + b + c;
  };
}
void main() {
  var step = makeChain(1)(2);
  print(step(3));
}"#,
        ["6"]
    };

    closure_stored_in_list_and_invoked => {
        r#"void main() {
  var fns = <int Function(int)>[
    (x) => x + 1,
    (x) => x * 2,
    (x) => x - 1,
  ];
  print(fns[0](5));
  print(fns[1](5));
  print(fns[2](5));
}"#,
        ["6", "10", "4"]
    };

    closure_immediately_invoked_reads_capture => {
        r#"void main() {
  var base = 8;
  print(((int n) => base + n)(2));
}"#,
        ["10"]
    };

    closure_passed_twice_with_different_captures => {
        r#"int run(int Function(int) fn, int x) {
  return fn(x);
}
void main() {
  var offset = 3;
  print(run((n) => n + offset, 10));
  offset = 5;
  print(run((n) => n + offset, 10));
}"#,
        ["13", "15"]
    };

    nested_closure_reads_two_outer_variables => {
        r#"void main() {
  var a = 2;
  var b = 3;
  var compute = () {
    var inner = () => a * b;
    return inner();
  };
  print(compute());
}"#,
        ["6"]
    };

    closure_parameter_shadows_captured_name => {
        r#"void main() {
  var x = 1;
  var fn = (x) => x + 10;
  print(fn(5));
  print(x);
}"#,
        ["15", "1"]
    };

    closure_reads_capture_after_outer_reassignment => {
        r#"void main() {
  var value = 1;
  var read = () => value;
  value = 9;
  print(read());
}"#,
        ["9"]
    };

    closure_used_as_where_predicate => {
        r#"void main() {
  var threshold = 3;
  var nums = [1, 2, 3, 4, 5];
  var filtered = nums.where((n) => n > threshold).toList();
  print(filtered.join(','));
}"#,
        ["4,5"]
    };

    closure_used_as_map_transform => {
        r#"void main() {
  var offset = 10;
  var nums = [1, 2, 3];
  var mapped = nums.map((n) => n + offset).toList();
  print(mapped.join('-'));
}"#,
        ["11-12-13"]
    };

    closure_compose_two_unary_functions => {
        r#"int compose(int x, int Function(int) f, int Function(int) g) {
  return f(g(x));
}
void main() {
  print(compose(3, (n) => n + 1, (n) => n * 2));
}"#,
        ["7"]
    };

    closure_as_custom_sort_comparator => {
        r#"void main() {
  var words = ['bb', 'a', 'ccc'];
  words.sort((a, b) => a.length.compareTo(b.length));
  print(words.join(','));
}"#,
        ["a,bb,ccc"]
    };

    closure_captures_mutable_flag_and_toggles => {
        r#"void main() {
  var on = false;
  var toggle = () {
    on = !on;
  };
  toggle();
  toggle();
  print(on);
}"#,
        ["false"]
    };

    closure_in_conditional_assignment => {
        r#"void main() {
  var useDouble = true;
  int Function(int) op = useDouble ? (x) => x * 2 : (x) => x + 1;
  print(op(5));
}"#,
        ["10"]
    };

    closure_called_multiple_times_reuses_capture => {
        r#"void main() {
  var base = 2;
  var addBase = (int n) => base + n;
  print(addBase(1));
  print(addBase(2));
  print(addBase(3));
}"#,
        ["3", "4", "5"]
    };

    closure_wrapper_delegates_to_captured_fn => {
        r#"void main() {
  int doubleIt(int x) => x * 2;
  var wrap = (int n) => doubleIt(n) + 1;
  print(wrap(4));
}"#,
        ["9"]
    };

    closure_factory_with_different_instances => {
        r#"int Function(int) makeOffset(int n) {
  return (x) => x - n;
}
void main() {
  var sub3 = makeOffset(3);
  var sub5 = makeOffset(5);
  print(sub3(10));
  print(sub5(10));
}"#,
        ["7", "5"]
    };

    closure_for_each_accumulates_with_capture => {
        r#"void main() {
  var total = 0;
  [1, 2, 3, 4].forEach((n) {
    total += n;
  });
  print(total);
}"#,
        ["10"]
    };

    closure_captures_map_and_updates_entry => {
        r#"void main() {
  var counts = <String, int>{'a': 1};
  var bump = (String key) {
    counts[key] = (counts[key] ?? 0) + 1;
  };
  bump('a');
  bump('b');
  print(counts['a']);
  print(counts['b']);
}"#,
        ["2", "1"]
    };

    closure_as_reduce_seed_combiner => {
        r#"void main() {
  var nums = [2, 3, 4];
  var product = nums.fold(1, (acc, n) => acc * n);
  print(product);
}"#,
        ["24"]
    };

    closure_with_two_parameters_as_argument => {
        r#"int combine(int a, int b, int Function(int, int) fn) {
  return fn(a, b);
}
void main() {
  print(combine(3, 4, (x, y) => x * y));
}"#,
        ["12"]
    };

    void_closure_runs_without_return_value => {
        r#"void runTwice(void Function() fn) {
  fn();
  fn();
}
void main() {
  var log = <String>[];
  runTwice(() {
    log.add('x');
  });
  print(log.length);
}"#,
        ["2"]
    };

    closure_captures_list_length_at_call_time => {
        r#"void main() {
  var items = [1, 2];
  var size = () => items.length;
  items.add(3);
  print(size());
}"#,
        ["3"]
    };
}
