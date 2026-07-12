//! Top-level functions, return semantics, recursion, local functions, and arrow functions.

dart_cases! {
    top_level_function_returns_sum_of_two_integers => {
        r#"int add(int a, int b) {
  return a + b;
}
void main() {
  print(add(3, 4));
}"#,
        ["7"]
    };

    top_level_function_returns_difference => {
        r#"int subtract(int a, int b) {
  return a - b;
}
void main() {
  print(subtract(10, 3));
}"#,
        ["7"]
    };

    top_level_function_returns_product => {
        r#"int multiply(int a, int b) {
  return a * b;
}
void main() {
  print(multiply(6, 7));
}"#,
        ["42"]
    };

    top_level_function_returns_quotient_truncated => {
        r#"int divide(int a, int b) {
  return a ~/ b;
}
void main() {
  print(divide(17, 5));
}"#,
        ["3"]
    };

    top_level_function_with_three_parameters => {
        r#"int sum3(int a, int b, int c) {
  return a + b + c;
}
void main() {
  print(sum3(1, 2, 3));
}"#,
        ["6"]
    };

    top_level_function_returns_string_concatenation => {
        r#"String greet(String name) {
  return 'Hello $name';
}
void main() {
  print(greet('Dart'));
}"#,
        ["Hello Dart"]
    };

    top_level_function_returns_boolean_comparison => {
        r#"bool isEven(int n) {
  return n % 2 == 0;
}
void main() {
  print(isEven(4));
  print(isEven(5));
}"#,
        ["true", "false"]
    };

    top_level_void_function_prints_side_effect => {
        r#"void shout(String msg) {
  print(msg);
}
void main() {
  shout('hi');
}"#,
        ["hi"]
    };

    top_level_function_returns_nullable_null => {
        r#"String? find(int id) {
  return null;
}
void main() {
  print(find(1));
}"#,
        ["null"]
    };

    top_level_function_returns_list_literal => {
        r#"List<int> makeRange(int n) {
  return [for (var i = 0; i < n; i++) i];
}
void main() {
  print(makeRange(4).join(','));
}"#,
        ["0,1,2,3"]
    };

    top_level_function_returns_map_literal => {
        r#"Map<String, int> scores() {
  return {'a': 1, 'b': 2};
}
void main() {
  print(scores()['a']);
  print(scores().length);
}"#,
        ["1", "2"]
    };

    return_statement_exits_function_early => {
        r#"int absDiff(int a, int b) {
  if (a >= b) {
    return a - b;
  }
  return b - a;
}
void main() {
  print(absDiff(5, 9));
  print(absDiff(9, 5));
}"#,
        ["4", "4"]
    };

    return_inside_if_branch_skips_else => {
        r#"String classify(int n) {
  if (n < 0) {
    return 'negative';
  }
  if (n == 0) {
    return 'zero';
  }
  return 'positive';
}
void main() {
  print(classify(-1));
  print(classify(0));
  print(classify(3));
}"#,
        ["negative", "zero", "positive"]
    };

    return_from_nested_block_inside_loop => {
        r#"int firstNegative(List<int> nums) {
  for (var n in nums) {
    if (n < 0) {
      return n;
    }
  }
  return 0;
}
void main() {
  print(firstNegative([1, 2, -3, 4]));
}"#,
        ["-3"]
    };

    void_function_explicit_return_exits => {
        r#"void logOnce(bool flag) {
  if (!flag) {
    return;
  }
  print('logged');
}
void main() {
  logOnce(false);
  print('after');
}"#,
        ["after"]
    };

    recursive_factorial_of_five => {
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

    recursive_factorial_base_case_zero => {
        r#"int fact(int n) {
  if (n <= 1) {
    return 1;
  }
  return n * fact(n - 1);
}
void main() {
  print(fact(0));
}"#,
        ["1"]
    };

    recursive_fibonacci_seventh => {
        r#"int fib(int n) {
  if (n <= 1) {
    return n;
  }
  return fib(n - 1) + fib(n - 2);
}
void main() {
  print(fib(7));
}"#,
        ["13"]
    };

    recursive_sum_one_through_ten => {
        r#"int sum(int n) {
  if (n <= 0) {
    return 0;
  }
  return n + sum(n - 1);
}
void main() {
  print(sum(10));
}"#,
        ["55"]
    };

    recursive_gcd_of_eighteen_and_twelve => {
        r#"int gcd(int a, int b) {
  if (b == 0) {
    return a;
  }
  return gcd(b, a % b);
}
void main() {
  print(gcd(18, 12));
}"#,
        ["6"]
    };

    recursive_power_two_to_fifth => {
        r#"int power(int base, int exp) {
  if (exp == 0) {
    return 1;
  }
  return base * power(base, exp - 1);
}
void main() {
  print(power(2, 5));
}"#,
        ["32"]
    };

    recursive_countdown_prints_to_one => {
        r#"void countdown(int n) {
  if (n == 0) {
    return;
  }
  print(n);
  countdown(n - 1);
}
void main() {
  countdown(3);
}"#,
        ["3", "2", "1"]
    };

    recursive_list_length_via_tail_style => {
        r#"int length(List<int> xs, int acc) {
  if (xs.isEmpty) {
    return acc;
  }
  return length(xs.sublist(1), acc + 1);
}
void main() {
  print(length([1, 2, 3, 4], 0));
}"#,
        ["4"]
    };

    local_function_adds_two_numbers => {
        r#"void main() {
  int add(int a, int b) {
    return a + b;
  }
  print(add(10, 20));
}"#,
        ["30"]
    };

    local_function_reads_outer_variable => {
        r#"void main() {
  var base = 100;
  int bump(int n) {
    return base + n;
  }
  print(bump(5));
}"#,
        ["105"]
    };

    local_function_calls_sibling_local_function => {
        r#"void main() {
  int double(int x) {
    return x * 2;
  }
  int quadruple(int x) {
    return double(double(x));
  }
  print(quadruple(3));
}"#,
        ["12"]
    };

    local_function_with_multiple_parameters => {
        r#"void main() {
  String label(int id, String name) {
    return '$id:$name';
  }
  print(label(7, 'vybe'));
}"#,
        ["7:vybe"]
    };

    local_recursive_function_computes_factorial => {
        r#"void main() {
  int fact(int n) {
    if (n <= 1) {
      return 1;
    }
    return n * fact(n - 1);
  }
  print(fact(6));
}"#,
        ["720"]
    };

    nested_local_functions_three_levels => {
        r#"void main() {
  int outer(int x) {
    int middle(int y) {
      int inner(int z) {
        return x + y + z;
      }
      return inner(3);
    }
    return middle(2);
  }
  print(outer(1));
}"#,
        ["6"]
    };

    local_function_shadows_outer_name => {
        r#"var value = 1;
int getValue() {
  return value;
}
void main() {
  int getValue() {
    return 99;
  }
  print(getValue());
}"#,
        ["99"]
    };

    top_level_arrow_function_squares_integer => {
        r#"int square(int x) => x * x;
void main() {
  print(square(8));
}"#,
        ["64"]
    };

    top_level_arrow_function_with_string_interpolation => {
        r#"String tag(String name) => 'tag:$name';
void main() {
  print(tag('core'));
}"#,
        ["tag:core"]
    };

    arrow_function_single_parameter_without_parens => {
        r#"void main() {
  var double = (x) => x * 2;
  print(double(11));
}"#,
        ["22"]
    };

    arrow_function_block_body_with_multiple_statements => {
        r#"void main() {
  var describe = (int n) {
    var sign = n < 0 ? 'neg' : 'pos';
    return '$sign:$n';
  };
  print(describe(-4));
  print(describe(4));
}"#,
        ["neg:-4", "pos:4"]
    };

    arrow_function_assigned_to_variable_and_called => {
        r#"void main() {
  bool Function(int) isPositive = (n) => n > 0;
  print(isPositive(1));
  print(isPositive(-1));
}"#,
        ["true", "false"]
    };

    arrow_function_used_as_foreach_callback => {
        r#"void main() {
  var sum = 0;
  [1, 2, 3].forEach((n) => sum += n);
  print(sum);
}"#,
        ["6"]
    };

    arrow_function_returns_string_from_block => {
        r#"void main() {
  String Function(int) fmt = (n) {
    return 'n=$n';
  };
  print(fmt(42));
}"#,
        ["n=42"]
    };

    arrow_function_void_body_runs_for_side_effect => {
        r#"void main() {
  var hits = 0;
  void Function() tick = () => hits++;
  tick();
  tick();
  print(hits);
}"#,
        ["2"]
    };

    local_arrow_function_inside_main => {
        r#"void main() {
  var triple = (int x) => x * 3;
  print(triple(5));
}"#,
        ["15"]
    };

    top_level_function_calls_other_top_level_function => {
        r#"int doubleIt(int x) {
  return x * 2;
}
int quadrupleIt(int x) {
  return doubleIt(doubleIt(x));
}
void main() {
  print(quadrupleIt(4));
}"#,
        ["16"]
    };

    function_reference_passed_as_first_class_value => {
        r#"int doubleIt(int x) => x * 2;
int applyTwice(int x, int Function(int) fn) {
  return fn(fn(x));
}
void main() {
  print(applyTwice(3, doubleIt));
}"#,
        ["12"]
    };

    function_returns_another_function => {
        r#"int Function(int) makeAdder(int n) {
  return (x) => x + n;
}
void main() {
  var add10 = makeAdder(10);
  print(add10(5));
}"#,
        ["15"]
    };

    top_level_function_with_default_path_after_loop => {
        r#"int indexOf(List<int> xs, int target) {
  for (var i = 0; i < xs.length; i++) {
    if (xs[i] == target) {
      return i;
    }
  }
  return -1;
}
void main() {
  print(indexOf([10, 20, 30], 20));
  print(indexOf([10, 20, 30], 99));
}"#,
        ["1", "-1"]
    };

    recursive_mutual_style_even_odd => {
        r#"bool isEven(int n) {
  if (n == 0) {
    return true;
  }
  return isOdd(n - 1);
}
bool isOdd(int n) {
  if (n == 0) {
    return false;
  }
  return isEven(n - 1);
}
void main() {
  print(isEven(4));
  print(isOdd(4));
}"#,
        ["true", "false"]
    };

    arrow_function_with_explicit_parameter_types => {
        r#"void main() {
  int add(int a, int b) => a + b;
  print(add(13, 29));
}"#,
        ["42"]
    };
}
