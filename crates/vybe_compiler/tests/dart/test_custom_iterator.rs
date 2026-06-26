//! Custom Iterator implementations: moveNext/current, Iterable with custom
//! iterator getter, manual while loops, and for-in over hand-built iterables.

dart_cases! {
    custom_iterator_move_next_returns_sequence => {
        r#"class ThreeIterator implements Iterator<int> {
  int _step = 0;
  @override
  int get current => _step;
  @override
  bool moveNext() {
    if (_step < 3) {
      _step = _step + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = ThreeIterator();
  var sum = 0;
  while (it.moveNext()) {
    sum = sum + it.current;
  }
  print(sum);
}"#,
        ["6"]
    };

    custom_iterator_manual_while_loop_counts => {
        r#"class CountIterator implements Iterator<int> {
  int _n = 0;
  final int limit;
  CountIterator(this.limit);
  @override
  int get current => _n;
  @override
  bool moveNext() {
    if (_n < limit) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = CountIterator(5);
  var count = 0;
  while (it.moveNext()) {
    count = count + 1;
  }
  print(count);
}"#,
        ["5"]
    };

    custom_iterator_empty_never_moves => {
        r#"class EmptyIterator implements Iterator<int> {
  @override
  int get current => 0;
  @override
  bool moveNext() => false;
}
void main() {
  var it = EmptyIterator();
  print(it.moveNext());
}"#,
        ["false"]
    };

    custom_iterator_single_element => {
        r#"class OnceIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 42;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void main() {
  var it = OnceIterator();
  it.moveNext();
  print(it.current);
}"#,
        ["42"]
    };

    custom_iterable_for_in_loop => {
        r#"class RangeIterable extends IterableBase<int> {
  final int start;
  final int end;
  RangeIterable(this.start, this.end);
  @override
  Iterator<int> get iterator => RangeIterator(start, end);
}
class RangeIterator implements Iterator<int> {
  int _current;
  final int end;
  RangeIterator(int start, int end) : _current = start - 1, end = end;
  @override
  int get current => _current;
  @override
  bool moveNext() {
    if (_current < end) {
      _current = _current + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var sum = 0;
  for (var n in RangeIterable(1, 4)) {
    sum = sum + n;
  }
  print(sum);
}"#,
        ["10"]
    };

    custom_iterable_iterator_getter_returns_new_instance => {
        r#"class RepeatOne extends IterableBase<int> {
  @override
  Iterator<int> get iterator => OneIterator();
}
class OneIterator implements Iterator<int> {
  bool _moved = false;
  @override
  int get current => 7;
  @override
  bool moveNext() {
    if (!_moved) {
      _moved = true;
      return true;
    }
    return false;
  }
}
void main() {
  var first = RepeatOne().iterator;
  var second = RepeatOne().iterator;
  first.moveNext();
  second.moveNext();
  print(first.current + second.current);
}"#,
        ["14"]
    };

    custom_iterator_two_independent_iterations => {
        r#"class TwoStep extends IterableBase<int> {
  @override
  Iterator<int> get iterator => StepIterator();
}
class StepIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var iterable = TwoStep();
  var a = 0;
  for (var n in iterable) {
    a = a + n;
  }
  var b = 0;
  for (var n in iterable) {
    b = b + n;
  }
  print(a);
  print(b);
}"#,
        ["3", "3"]
    };

    custom_iterator_string_values => {
        r#"class WordIterator implements Iterator<String> {
  final List<String> words;
  int _i = -1;
  WordIterator(this.words);
  @override
  String get current => words[_i];
  @override
  bool moveNext() {
    if (_i + 1 < words.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = WordIterator(['a', 'b', 'c']);
  var text = '';
  while (it.moveNext()) {
    text = text + it.current;
  }
  print(text);
}"#,
        ["abc"]
    };

    custom_iterator_even_numbers_only => {
        r#"class Evens extends IterableBase<int> {
  final int max;
  Evens(this.max);
  @override
  Iterator<int> get iterator => EvensIterator(max);
}
class EvensIterator implements Iterator<int> {
  int _n = 0;
  final int max;
  EvensIterator(this.max);
  @override
  int get current => _n;
  @override
  bool moveNext() {
  _n = _n + 2;
    if (_n <= max) {
      return true;
    }
    return false;
  }
}
void main() {
  var list = <int>[];
  for (var n in Evens(6)) {
    list.add(n);
  }
  print(list.join(','));
}"#,
        ["2,4,6"]
    };

    custom_iterator_countdown_sequence => {
        r#"class Countdown extends IterableBase<int> {
  final int start;
  Countdown(this.start);
  @override
  Iterator<int> get iterator => CountdownIterator(start);
}
class CountdownIterator implements Iterator<int> {
  int _n;
  int _current = 0;
  CountdownIterator(int start) : _n = start;
  @override
  int get current => _current;
  @override
  bool moveNext() {
    if (_n > 0) {
      _current = _n;
      _n = _n - 1;
      return true;
    }
    return false;
  }
}
void main() {
  var text = '';
  for (var n in Countdown(3)) {
    text = text + n.toString();
  }
  print(text);
}"#,
        ["321"]
    };

    custom_iterable_to_list_conversion => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().toList().join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterable_join_elements => {
        r#"class Letters extends IterableBase<String> {
  @override
  Iterator<String> get iterator => LetterIterator();
}
class LetterIterator implements Iterator<String> {
  int _i = -1;
  final codes = ['x', 'y', 'z'];
  @override
  String get current => codes[_i];
  @override
  bool moveNext() {
    if (_i + 1 < codes.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Letters().join('-'));
}"#,
        ["x-y-z"]
    };

    custom_iterable_length_property => {
        r#"class Fixed extends IterableBase<int> {
  final int count;
  Fixed(this.count);
  @override
  Iterator<int> get iterator => FixedIterator(count);
  @override
  int get length => count;
}
class FixedIterator implements Iterator<int> {
  int _n = 0;
  final int count;
  FixedIterator(this.count);
  @override
  int get current => _n;
  @override
  bool moveNext() {
    if (_n < count) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Fixed(4).length);
}"#,
        ["4"]
    };

    custom_iterable_is_empty_false => {
        r#"class OneItem extends IterableBase<int> {
  @override
  Iterator<int> get iterator => OneItemIterator();
}
class OneItemIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 1;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void main() {
  print(OneItem().isEmpty);
}"#,
        ["false"]
    };

    custom_iterable_is_not_empty => {
        r#"class OneItem extends IterableBase<int> {
  @override
  Iterator<int> get iterator => OneItemIterator();
}
class OneItemIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 1;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void main() {
  print(OneItem().isNotEmpty);
}"#,
        ["true"]
    };

    custom_iterable_contains_check => {
        r#"class Small extends IterableBase<int> {
  @override
  Iterator<int> get iterator => SmallIterator();
}
class SmallIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Small().contains(2));
}"#,
        ["true"]
    };

    custom_iterable_first_element => {
        r#"class FirstFive extends IterableBase<int> {
  @override
  Iterator<int> get iterator => FirstFiveIterator();
}
class FirstFiveIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(FirstFive().first);
}"#,
        ["1"]
    };

    custom_iterable_last_element => {
        r#"class FirstFive extends IterableBase<int> {
  @override
  Iterator<int> get iterator => FirstFiveIterator();
}
class FirstFiveIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(FirstFive().last);
}"#,
        ["5"]
    };

    custom_iterable_element_at_index => {
        r#"class Indexed extends IterableBase<int> {
  @override
  Iterator<int> get iterator => IndexedIterator();
}
class IndexedIterator implements Iterator<int> {
  int _v = -1;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Indexed().elementAt(2));
}"#,
        ["2"]
    };

    custom_iterable_map_transformation => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().map((n) => n * 2).join(','));
}"#,
        ["2,4,6"]
    };

    custom_iterable_where_filter => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().where((n) => n % 2 == 0).join(','));
}"#,
        ["2,4"]
    };

    custom_iterable_skip_prefix => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 4) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().skip(2).join(','));
}"#,
        ["3,4"]
    };

    custom_iterable_take_prefix => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().take(2).join(','));
}"#,
        ["1,2"]
    };

    custom_iterable_fold_accumulator => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 4) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().fold(0, (acc, n) => acc + n));
}"#,
        ["10"]
    };

    custom_iterable_reduce_combines => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().reduce((a, b) => a * b));
}"#,
        ["6"]
    };

    custom_iterable_every_predicate => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().every((n) => n > 0));
}"#,
        ["true"]
    };

    custom_iterable_any_predicate => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().any((n) => n == 4));
}"#,
        ["true"]
    };

    custom_iterator_for_in_with_break => {
        r#"class Long extends IterableBase<int> {
  @override
  Iterator<int> get iterator => LongIterator();
}
class LongIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 10) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var sum = 0;
  for (var n in Long()) {
    if (n == 3) {
      break;
    }
    sum = sum + n;
  }
  print(sum);
}"#,
        ["3"]
    };

    custom_iterator_for_in_with_continue => {
        r#"class Long extends IterableBase<int> {
  @override
  Iterator<int> get iterator => LongIterator();
}
class LongIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var sum = 0;
  for (var n in Long()) {
    if (n == 2) {
      continue;
    }
    sum = sum + n;
  }
  print(sum);
}"#,
        ["13"]
    };

    custom_iterator_move_next_side_effect => {
        r#"class LoggingIterator implements Iterator<int> {
  int _v = 0;
  int moves = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    moves = moves + 1;
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = LoggingIterator();
  while (it.moveNext()) {}
  print(it.moves);
}"#,
        ["3"]
    };

    custom_iterable_fibonacci_limited => {
        r#"class Fib extends IterableBase<int> {
  final int count;
  Fib(this.count);
  @override
  Iterator<int> get iterator => FibIterator(count);
}
class FibIterator implements Iterator<int> {
  int _a = 0;
  int _b = 1;
  int _seen = 0;
  final int count;
  FibIterator(this.count);
  @override
  int get current => _a;
  @override
  bool moveNext() {
    if (_seen >= count) {
      return false;
    }
    var next = _a + _b;
    _a = _b;
    _b = next;
    _seen = _seen + 1;
    return true;
  }
}
void main() {
  print(Fib(5).join(','));
}"#,
        ["1,1,2,3,5"]
    };

    custom_iterable_modulo_pattern => {
        r#"class ModThree extends IterableBase<int> {
  @override
  Iterator<int> get iterator => ModThreeIterator();
}
class ModThreeIterator implements Iterator<int> {
  int _n = 0;
  @override
  int get current => _n % 3;
  @override
  bool moveNext() {
    if (_n < 6) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(ModThree().join(','));
}"#,
        ["1,2,0,1,2,0"]
    };

    custom_iterable_doubles_sequence => {
        r#"class Halves extends IterableBase<double> {
  @override
  Iterator<double> get iterator => HalvesIterator();
}
class HalvesIterator implements Iterator<double> {
  int _n = 0;
  @override
  double get current => _n * 0.5;
  @override
  bool moveNext() {
    if (_n < 4) {
      _n = _n + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Halves().map((d) => d.toString()).join(','));
}"#,
        ["0.5,1.0,1.5,2.0"]
    };

    custom_iterable_subclass_extends_base => {
        r#"class BaseRange extends IterableBase<int> {
  final int n;
  BaseRange(this.n);
  @override
  Iterator<int> get iterator => BaseRangeIterator(n);
}
class BaseRangeIterator implements Iterator<int> {
  int _v = 0;
  final int n;
  BaseRangeIterator(this.n);
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < n) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
class DoubleRange extends BaseRange {
  DoubleRange(int n) : super(n);
}
void main() {
  print(DoubleRange(3).join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterator_indexed_manual_loop => {
        r#"class IndexWalk extends IterableBase<int> {
  @override
  Iterator<int> get iterator => IndexWalkIterator();
}
class IndexWalkIterator implements Iterator<int> {
  int _i = -1;
  @override
  int get current => _i;
  @override
  bool moveNext() {
    if (_i < 2) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = IndexWalk().iterator;
  var idx = 0;
  while (it.moveNext()) {
    print(it.current);
    idx = idx + 1;
  }
  print(idx);
}"#,
        ["0", "1", "2", "3"]
    };

    custom_iterable_followed_by_chain => {
        r#"class ABC extends IterableBase<String> {
  @override
  Iterator<String> get iterator => ABCIterator();
}
class ABCIterator implements Iterator<String> {
  int _i = -1;
  final items = ['a', 'b'];
  @override
  String get current => items[_i];
  @override
  bool moveNext() {
    if (_i + 1 < items.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(ABC().followedBy(['c']).join(''));
}"#,
        ["abc"]
    };

    custom_iterable_expand_flatten => {
        r#"class Pairs extends IterableBase<List<int>> {
  @override
  Iterator<List<int>> get iterator => PairsIterator();
}
class PairsIterator implements Iterator<List<int>> {
  int _step = 0;
  @override
  List<int> get current => _step == 0 ? [1, 2] : [3];
  @override
  bool moveNext() {
    if (_step < 2) {
      _step = _step + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Pairs().expand((p) => p).join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterable_cast_typed_view => {
        r#"class AnyNums extends IterableBase<num> {
  @override
  Iterator<num> get iterator => AnyNumsIterator();
}
class AnyNumsIterator implements Iterator<num> {
  int _v = 0;
  @override
  num get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(AnyNums().cast<int>().join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterable_to_set_unique => {
        r#"class Dupes extends IterableBase<int> {
  @override
  Iterator<int> get iterator => DupesIterator();
}
class DupesIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v <= 2 ? 1 : 2;
  @override
  bool moveNext() {
    if (_v < 4) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Dupes().toSet().length);
}"#,
        ["2"]
    };

    custom_iterator_triple_manual_unroll => {
        r#"class TripleIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 3) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = TripleIterator();
  var parts = <int>[];
  if (it.moveNext()) {
    parts.add(it.current);
  }
  if (it.moveNext()) {
    parts.add(it.current);
  }
  if (it.moveNext()) {
    parts.add(it.current);
  }
  print(parts.join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterable_empty_is_empty => {
        r#"class Nothing extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NothingIterator();
}
class NothingIterator implements Iterator<int> {
  @override
  int get current => 0;
  @override
  bool moveNext() => false;
}
void main() {
  print(Nothing().isEmpty);
}"#,
        ["true"]
    };

    custom_iterable_single_length_one => {
        r#"class Single extends IterableBase<int> {
  @override
  Iterator<int> get iterator => SingleIterator();
  @override
  int get length => 1;
}
class SingleIterator implements Iterator<int> {
  bool _done = false;
  @override
  int get current => 9;
  @override
  bool moveNext() {
    if (!_done) {
      _done = true;
      return true;
    }
    return false;
  }
}
void main() {
  print(Single().length);
}"#,
        ["1"]
    };

    custom_iterator_powers_of_two => {
        r#"class Powers extends IterableBase<int> {
  @override
  Iterator<int> get iterator => PowersIterator();
}
class PowersIterator implements Iterator<int> {
  final values = [1, 2, 4, 8];
  int _i = -1;
  @override
  int get current => values[_i];
  @override
  bool moveNext() {
    if (_i + 1 < values.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Powers().join(','));
}"#,
        ["1,2,4,8"]
    };

    custom_iterable_indexed_for_loop_equivalent => {
        r#"class ZeroToTwo extends IterableBase<int> {
  @override
  Iterator<int> get iterator => ZeroToTwoIterator();
}
class ZeroToTwoIterator implements Iterator<int> {
  int _v = -1;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 2) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var total = 0;
  for (var i = 0; i < 3; i++) {
    total = total + i;
  }
  var iterTotal = 0;
  for (var n in ZeroToTwo()) {
    iterTotal = iterTotal + n;
  }
  print(total);
  print(iterTotal);
}"#,
        ["3", "3"]
    };

    custom_iterator_alternating_bool_move_next => {
        r#"class AltIterator implements Iterator<int> {
  int _count = 0;
  @override
  int get current => _count;
  @override
  bool moveNext() {
    if (_count < 3) {
      _count = _count + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = AltIterator();
  print(it.moveNext());
  print(it.moveNext());
  print(it.moveNext());
  print(it.moveNext());
}"#,
        ["true", "true", "true", "false"]
    };

    custom_iterable_reversed_manual_build => {
        r#"class Rev extends IterableBase<int> {
  final List<int> data;
  Rev(this.data);
  @override
  Iterator<int> get iterator => RevIterator(data);
}
class RevIterator implements Iterator<int> {
  int _i;
  final List<int> data;
  RevIterator(List<int> data) : data = data, _i = data.length;
  @override
  int get current => data[_i];
  @override
  bool moveNext() {
    if (_i > 0) {
      _i = _i - 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Rev([1, 2, 3]).join(','));
}"#,
        ["3,2,1"]
    };

    custom_iterable_skip_while_tail => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().skipWhile((n) => n < 3).join(','));
}"#,
        ["3,4,5"]
    };

    custom_iterable_take_while_prefix => {
        r#"class Nums extends IterableBase<int> {
  @override
  Iterator<int> get iterator => NumsIterator();
}
class NumsIterator implements Iterator<int> {
  int _v = 0;
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < 5) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Nums().takeWhile((n) => n < 4).join(','));
}"#,
        ["1,2,3"]
    };

    custom_iterator_while_loop_accumulates_product => {
        r#"class MultIterator implements Iterator<int> {
  int _v = 0;
  final int limit;
  MultIterator(this.limit);
  @override
  int get current => _v;
  @override
  bool moveNext() {
    if (_v < limit) {
      _v = _v + 1;
      return true;
    }
    return false;
  }
}
void main() {
  var it = MultIterator(3);
  var product = 1;
  while (it.moveNext()) {
    product = product * it.current;
  }
  print(product);
}"#,
        ["6"]
    };

    custom_iterable_enumerate_style_manual => {
        r#"class Tagged extends IterableBase<String> {
  final List<String> items;
  Tagged(this.items);
  @override
  Iterator<String> get iterator => TaggedIterator(items);
}
class TaggedIterator implements Iterator<String> {
  int _i = -1;
  final List<String> items;
  TaggedIterator(this.items);
  @override
  String get current => '${_i}:${items[_i]}';
  @override
  bool moveNext() {
    if (_i + 1 < items.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void main() {
  print(Tagged(['a', 'b']).join('|'));
}"#,
        ["0:a|1:b"]
    };
}
