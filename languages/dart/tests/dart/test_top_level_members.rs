//! Top-level variables, functions, getters, setters, const values,
//! and cross-function interaction at library scope.

dart_cases! {
    top_level_var_mutable_reassignment => {
        r#"var counter = 0;
void main() {
  counter = counter + 1;
  print(counter);
}"#,
        ["1"]
    };

    top_level_final_initialized_once => {
        r#"final maxRetries = 3;
void main() {
  print(maxRetries);
}"#,
        ["3"]
    };

    top_level_const_compile_time_value => {
        r#"const piApprox = 314;
void main() {
  print(piApprox ~/ 100);
}"#,
        ["3"]
    };

    top_level_typed_int_variable => {
        r#"int total = 100;
void main() {
  print(total - 40);
}"#,
        ["60"]
    };

    top_level_typed_string_variable => {
        r#"String greeting = 'hello';
void main() {
  print(greeting.length);
}"#,
        ["5"]
    };

    top_level_bool_flag => {
        r#"bool enabled = true;
void main() {
  print(enabled);
}"#,
        ["true"]
    };

    top_level_list_literal_mutable => {
        r#"var items = [1, 2, 3];
void main() {
  items.add(4);
  print(items.length);
}"#,
        ["4"]
    };

    top_level_map_literal_access => {
        r#"var config = {'port': 8080};
void main() {
  print(config['port']);
}"#,
        ["8080"]
    };

    top_level_function_adds_two_integers => {
        r#"int add(int a, int b) {
  return a + b;
}
void main() {
  print(add(10, 5));
}"#,
        ["15"]
    };

    top_level_function_returns_string => {
        r#"String label(int id) {
  return 'item-$id';
}
void main() {
  print(label(7));
}"#,
        ["item-7"]
    };

    top_level_void_function_with_side_effect => {
        r#"void emit(String msg) {
  print(msg);
}
void main() {
  emit('ping');
}"#,
        ["ping"]
    };

    top_level_getter_computes_from_var => {
        r#"int base = 10;
int get doubled {
  return base * 2;
}
void main() {
  print(doubled);
}"#,
        ["20"]
    };

    top_level_setter_updates_backing_var => {
        r#"int _score = 0;
int get score {
  return _score;
}
set score(int v) {
  _score = v;
}
void main() {
  score = 42;
  print(score);
}"#,
        ["42"]
    };

    top_level_getter_with_no_setter_read_only => {
        r#"int secret = 99;
int get masked {
  return secret ~/ 10;
}
void main() {
  print(masked);
}"#,
        ["9"]
    };

    top_level_const_used_in_function => {
        r#"const factor = 5;
int scale(int n) {
  return n * factor;
}
void main() {
  print(scale(6));
}"#,
        ["30"]
    };

    top_level_functions_call_each_other => {
        r#"int stepA(int n) {
  return n + 1;
}
int stepB(int n) {
  return stepA(n) * 2;
}
void main() {
  print(stepB(4));
}"#,
        ["10"]
    };

    top_level_chain_three_functions => {
        r#"int inc(int n) {
  return n + 1;
}
int twice(int n) {
  return inc(n) + inc(n);
}
int run(int n) {
  return twice(n);
}
void main() {
  print(run(3));
}"#,
        ["8"]
    };

    top_level_var_read_by_multiple_functions => {
        r#"int shared = 5;
int readShared() {
  return shared;
}
int bumpShared() {
  shared = shared + 1;
  return shared;
}
void main() {
  print(readShared());
  print(bumpShared());
  print(readShared());
}"#,
        ["5", "6", "6"]
    };

    top_level_function_mutates_top_level_var => {
        r#"int tally = 0;
void bump() {
  tally = tally + 1;
}
void main() {
  bump();
  bump();
  print(tally);
}"#,
        ["2"]
    };

    top_level_getter_reflects_var_mutation => {
        r#"int units = 1;
int get totalUnits {
  return units;
}
void main() {
  units = 4;
  print(totalUnits);
}"#,
        ["4"]
    };

    top_level_const_string_concat => {
        r#"const prefix = 'vybe';
const suffix = 'vm';
const name = prefix + suffix;
void main() {
  print(name);
}"#,
        ["vybevm"]
    };

    top_level_late_var_initialized_before_use => {
        r#"late int ready;
void prime() {
  ready = 7;
}
void main() {
  prime();
  print(ready);
}"#,
        ["7"]
    };

    top_level_function_with_default_return_path => {
        r#"int sign(int n) {
  if (n > 0) {
    return 1;
  }
  if (n < 0) {
    return -1;
  }
  return 0;
}
void main() {
  print(sign(0));
}"#,
        ["0"]
    };

    top_level_recursive_function_factorial => {
        r#"int fact(int n) {
  if (n <= 1) {
    return 1;
  }
  return n * fact(n - 1);
}
void main() {
  print(fact(5));
}"#,
        ["120"]
    };

    top_level_function_returns_list => {
        r#"List<int> range(int n) {
  return [for (var i = 0; i < n; i++) i];
}
void main() {
  print(range(4).join(','));
}"#,
        ["0,1,2,3"]
    };

    top_level_function_returns_map => {
        r#"Map<String, int> scores() {
  return {'a': 1, 'b': 2};
}
void main() {
  print(scores()['b']);
}"#,
        ["2"]
    };

    top_level_getter_setter_pair_for_counter => {
        r#"int _hits = 0;
int get hits {
  return _hits;
}
set hits(int v) {
  _hits = v;
}
void main() {
  hits = hits + 3;
  print(hits);
}"#,
        ["3"]
    };

    top_level_const_list_length => {
        r#"const primes = [2, 3, 5];
void main() {
  print(primes.length);
}"#,
        ["3"]
    };

    top_level_const_map_lookup => {
        r#"const codes = {'ok': 200, 'err': 500};
void main() {
  print(codes['ok']);
}"#,
        ["200"]
    };

    top_level_functions_mutual_no_cycle => {
        r#"int toA(int n) {
  return n + 10;
}
int toB(int n) {
  return toA(n) + 5;
}
void main() {
  print(toB(1));
}"#,
        ["16"]
    };

    top_level_var_nullable_defaults_null => {
        r#"String? nickname;
void main() {
  print(nickname);
}"#,
        ["null"]
    };

    top_level_function_accepts_nullable => {
        r#"String show(String? msg) {
  return msg ?? 'none';
}
void main() {
  print(show(null));
}"#,
        ["none"]
    };

    top_level_arrow_function => {
        r#"int square(int n) => n * n;
void main() {
  print(square(8));
}"#,
        ["64"]
    };

    top_level_getter_arrow_form => {
        r#"int seed = 6;
int get triple => seed * 3;
void main() {
  print(triple);
}"#,
        ["18"]
    };

    top_level_multiple_vars_independent => {
        r#"int a = 1;
int b = 2;
void main() {
  print(a + b);
}"#,
        ["3"]
    };

    top_level_function_uses_top_level_getter => {
        r#"int base = 5;
int get offset {
  return 2;
}
int compute() {
  return base + offset;
}
void main() {
  print(compute());
}"#,
        ["7"]
    };

    top_level_setter_called_from_function => {
        r#"int _level = 0;
set level(int v) {
  _level = v;
}
int get level {
  return _level;
}
void setLevel(int v) {
  level = v;
}
void main() {
  setLevel(9);
  print(level);
}"#,
        ["9"]
    };

    top_level_const_bool_flag => {
        r#"const debug = true;
void main() {
  print(debug);
}"#,
        ["true"]
    };

    top_level_function_closure_over_var => {
        r#"int factor = 3;
int apply(int n) {
  return n * factor;
}
void main() {
  factor = 4;
  print(apply(5));
}"#,
        ["20"]
    };

    top_level_pipeline_three_calls => {
        r#"int parse(String s) {
  return int.parse(s);
}
int doubleIt(int n) {
  return n * 2;
}
int finish(int n) {
  return n + 1;
}
void main() {
  print(finish(doubleIt(parse('10'))));
}"#,
        ["21"]
    };

    top_level_var_double_type => {
        r#"double ratio = 2.5;
void main() {
  print(ratio + 0.5);
}"#,
        ["3.0"]
    };

    top_level_function_with_named_params => {
        r#"int combine({int a = 1, int b = 2}) {
  return a + b;
}
void main() {
  print(combine(a: 10, b: 20));
}"#,
        ["30"]
    };

    top_level_getter_string_interpolation => {
        r#"String host = 'localhost';
String get endpoint {
  return 'http://$host';
}
void main() {
  print(endpoint.contains('localhost'));
}"#,
        ["true"]
    };

    top_level_const_used_by_getter => {
        r#"const unit = 10;
int get tens {
  return unit;
}
void main() {
  print(tens);
}"#,
        ["10"]
    };

    top_level_init_order_var_before_function_use => {
        r#"int limit = 5;
bool within(int n) {
  return n < limit;
}
void main() {
  print(within(3));
  print(within(9));
}"#,
        ["true", "false"]
    };
}
