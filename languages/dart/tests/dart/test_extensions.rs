//! Dart extension methods on int, String, List, and Iterable — instance
//! methods, getters, and static extension helpers.

dart_cases! {
    // ── extension on a USER-DECLARED class ───────────────────────────
    //
    // Distinct mechanism from every other case in this file. Extending a type
    // you OWN is member augmentation — the same operation as a mixin or a PHP
    // trait, and it folds into the class. Extending a BUILT-IN (every other
    // case below) cannot fold, because the type is not yours; that resolves
    // through the receiver's prototype instead. Two mechanisms, one syntax.
    // See flexclassplan.md §4c / §4d.

    extension_on_user_class_method => {
        r#"class Box { int v = 3; }
extension BoxX on Box { int twice() => v * 2; }
void main() {
  print(Box().twice());
}"#,
        ["6"]
    };

    extension_on_user_class_getter => {
        r#"class Box { int v = 4; }
extension BoxX on Box { int get tripled => v * 3; }
void main() {
  print(Box().tripled);
}"#,
        ["12"]
    };

    extension_on_user_class_does_not_override_own_member => {
        r#"class Box {
  int v = 5;
  int twice() => 999;
}
extension BoxX on Box { int twice() => v * 2; }
void main() {
  print(Box().twice());
}"#,
        ["999"]
    };

    // ── int: static extension methods ────────────────────────────────

    int_static_max_of_two_values => {
        r#"extension IntMath on int {
  static int max(int a, int b) => a > b ? a : b;
}
void main() {
  print(IntMath.max(10, 25));
}"#,
        ["25"]
    };

    int_static_min_of_two_values => {
        r#"extension IntMath on int {
  static int minVal(int a, int b) => a < b ? a : b;
}
void main() {
  print(IntMath.minVal(10, 25));
}"#,
        ["10"]
    };

    int_static_parse_decimal_string => {
        r#"extension IntParse on int {
  static int parseDec(String s) => int.parse(s);
}
void main() {
  print(IntParse.parseDec('456'));
}"#,
        ["456"]
    };

    int_static_sum_three_integers => {
        r#"extension IntAgg on int {
  static int sum3(int a, int b, int c) => a + b + c;
}
void main() {
  print(IntAgg.sum3(4, 5, 6));
}"#,
        ["15"]
    };

    int_static_absolute_difference => {
        r#"extension IntDiff on int {
  static int absDiff(int a, int b) => a > b ? a - b : b - a;
}
void main() {
  print(IntDiff.absDiff(5, 12));
}"#,
        ["7"]
    };

    int_static_clamp_between_bounds => {
        r#"extension IntClamp on int {
  static int clampVal(int n, int lo, int hi) {
    if (n < lo) return lo;
    if (n > hi) return hi;
    return n;
  }
}
void main() {
  print(IntClamp.clampVal(15, 0, 10));
}"#,
        ["10"]
    };

    int_static_is_within_inclusive_range => {
        r#"extension IntRange on int {
  static bool inRange(int n, int lo, int hi) => n >= lo && n <= hi;
}
void main() {
  print(IntRange.inRange(7, 1, 10));
}"#,
        ["true"]
    };

    // ── int: getters ─────────────────────────────────────────────────

    int_getter_is_odd_on_seven => {
        r#"extension IntParity on int {
  bool get isOdd => this % 2 != 0;
}
void main() {
  print(7.isOdd);
}"#,
        ["true"]
    };

    int_getter_is_even_on_four => {
        r#"extension IntParity on int {
  bool get isEven => this % 2 == 0;
}
void main() {
  print(4.isEven);
}"#,
        ["true"]
    };

    int_getter_sign_label_negative => {
        r#"extension IntSign on int {
  String get signLabel => this >= 0 ? 'pos' : 'neg';
}
void main() {
  print((-3).signLabel);
}"#,
        ["neg"]
    };

    // ── int: instance methods ────────────────────────────────────────

    int_method_doubles_value => {
        r#"extension IntTwice on int {
  int doubled() => this * 2;
}
void main() {
  print(5.doubled());
}"#,
        ["10"]
    };

    int_method_squares_value => {
        r#"extension IntSquare on int {
  int squared() => this * this;
}
void main() {
  print(7.squared());
}"#,
        ["49"]
    };

    int_method_times_multiplier => {
        r#"extension IntScale on int {
  int times(int factor) => this * factor;
}
void main() {
  print(3.times(4));
}"#,
        ["12"]
    };

    // ── String: static extension methods ─────────────────────────────

    string_static_length_of_text => {
        r#"extension StrLen on String {
  static int lengthOf(String s) => s.length;
}
void main() {
  print(StrLen.lengthOf('dart'));
}"#,
        ["4"]
    };

    string_static_concat_two_strings => {
        r#"extension StrJoin on String {
  static String concat(String a, String b) => a + b;
}
void main() {
  print(StrJoin.concat('foo', 'bar'));
}"#,
        ["foobar"]
    };

    string_static_starts_with_prefix => {
        r#"extension StrPrefix on String {
  static bool starts(String s, String prefix) => s.startsWith(prefix);
}
void main() {
  print(StrPrefix.starts('hello', 'he'));
}"#,
        ["true"]
    };

    // ── String: getters ──────────────────────────────────────────────

    string_getter_is_all_uppercase => {
        r#"extension StrCase on String {
  bool get isAllUpper => this == toUpperCase();
}
void main() {
  print('ABC'.isAllUpper);
}"#,
        ["true"]
    };

    string_getter_is_blank_whitespace => {
        r#"extension StrBlank on String {
  bool get isBlank => trim().isEmpty;
}
void main() {
  print('   '.isBlank);
}"#,
        ["true"]
    };

    string_getter_word_count_by_spaces => {
        r#"extension StrWords on String {
  int get wordCount => trim().split(' ').length;
}
void main() {
  print('one two three'.wordCount);
}"#,
        ["3"]
    };

    // ── String: instance methods ─────────────────────────────────────

    string_method_wrap_in_brackets => {
        r#"extension StrWrap on String {
  String wrap() => '<' + this + '>';
}
void main() {
  print('hi'.wrap());
}"#,
        ["<hi>"]
    };

    string_method_repeat_text => {
        r#"extension StrRepeat on String {
  String repeatText(int n) {
    var out = '';
    for (var i = 0; i < n; i++) {
      out += this;
    }
    return out;
  }
}
void main() {
  print('a'.repeatText(2));
}"#,
        ["aa"]
    };

    string_method_capitalize_first_letter => {
        r#"extension StrCap on String {
  String capitalize() {
    if (isEmpty) return this;
    return this[0].toUpperCase() + substring(1);
  }
}
void main() {
  print('hello'.capitalize());
}"#,
        ["Hello"]
    };

    string_method_remove_all_spaces => {
        r#"extension StrSpace on String {
  String removeSpaces() => split(' ').join('');
}
void main() {
  print('a b c'.removeSpaces());
}"#,
        ["abc"]
    };

    // ── List: static extension methods ───────────────────────────────

    list_static_sum_int_elements => {
        r#"extension ListSum on List<int> {
  static int total(List<int> nums) {
    var sum = 0;
    for (var n in nums) {
      sum += n;
    }
    return sum;
  }
}
void main() {
  print(ListSum.total([1, 2, 3, 4]));
}"#,
        ["10"]
    };

    list_static_count_elements => {
        r#"extension ListCount on List<int> {
  static int size(List<int> list) => list.length;
}
void main() {
  print(ListCount.size([10, 20, 30]));
}"#,
        ["3"]
    };

    // ── List: getters ────────────────────────────────────────────────

    list_getter_first_element_value => {
        r#"extension ListHead on List<int> {
  int get firstVal => this[0];
}
void main() {
  print([9, 8, 7].firstVal);
}"#,
        ["9"]
    };

    list_getter_last_element_value => {
        r#"extension ListTail on List<int> {
  int get lastVal => this[length - 1];
}
void main() {
  print([1, 2, 3].lastVal);
}"#,
        ["3"]
    };

    list_getter_is_singleton_list => {
        r#"extension ListSingle on List<int> {
  bool get isSingleton => length == 1;
}
void main() {
  print([42].isSingleton);
}"#,
        ["true"]
    };

    // ── List: instance methods ───────────────────────────────────────

    list_method_join_with_dash => {
        r#"extension ListJoin on List<int> {
  String joinDash() => map((n) => n.toString()).join('-');
}
void main() {
  print([1, 2, 3].joinDash());
}"#,
        ["1-2-3"]
    };

    list_method_contains_all_values => {
        r#"extension ListHas on List<int> {
  bool containsAllExt(List<int> others) {
    for (var v in others) {
      if (!contains(v)) return false;
    }
    return true;
  }
}
void main() {
  print([1, 2, 3, 4].containsAllExt([2, 4]));
}"#,
        ["true"]
    };

    list_method_copy_reversed_order => {
        r#"extension ListFlip on List<int> {
  List<int> reversedCopy() => reversed.toList();
}
void main() {
  print([1, 2, 3].reversedCopy().join(','));
}"#,
        ["3,2,1"]
    };

    // ── Iterable: static extension methods ───────────────────────────

    iterable_static_count_items => {
        r#"extension IterSize on Iterable<int> {
  static int countItems(Iterable<int> it) {
    var c = 0;
    for (var _ in it) {
      c++;
    }
    return c;
  }
}
void main() {
  print(IterSize.countItems([10, 20]));
}"#,
        ["2"]
    };

    // ── Iterable: getters ────────────────────────────────────────────

    iterable_getter_has_any_item => {
        r#"extension IterAny on Iterable<int> {
  bool get hasItems => !isEmpty;
}
void main() {
  print([1].hasItems);
}"#,
        ["true"]
    };

    iterable_getter_any_value_is_positive => {
        r#"extension IterPos on Iterable<int> {
  bool get anyPositive => any((n) => n > 0);
}
void main() {
  print([-1, 0, 2].anyPositive);
}"#,
        ["true"]
    };

    iterable_getter_all_elements_negative => {
        r#"extension IterNeg on Iterable<int> {
  bool get allNegative => every((n) => n < 0);
}
void main() {
  print([-1, -2].allNegative);
}"#,
        ["true"]
    };

    // ── Iterable: instance methods ───────────────────────────────────

    iterable_method_join_with_plus => {
        r#"extension IterJoin on Iterable<int> {
  String joinPlus() => map((n) => n.toString()).join('+');
}
void main() {
  print([2, 3].joinPlus());
}"#,
        ["2+3"]
    };

    iterable_method_first_matching_or_zero => {
        r#"extension IterFind on Iterable<int> {
  int firstGreaterThan(int threshold) {
    for (var n in this) {
      if (n > threshold) return n;
    }
    return 0;
  }
}
void main() {
  print([1, 5, 3].firstGreaterThan(2));
}"#,
        ["5"]
    };
}
