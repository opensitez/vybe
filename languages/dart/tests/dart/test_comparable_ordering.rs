//! Comparable ordering: compareTo, sort, and comparison operators via compareTo.

dart_cases! {
    comparable_compare_equal_returns_zero => {
        r#"class Rank implements Comparable<Rank> {
  int level;
  Rank(this.level);
  int compareTo(Rank other) => level.compareTo(other.level);
}
void main() {
  print(Rank(5).compareTo(Rank(5)));
}"#,
        ["0"]
    };

    comparable_compare_less_returns_negative => {
        r#"class Rank implements Comparable<Rank> {
  int level;
  Rank(this.level);
  int compareTo(Rank other) => level.compareTo(other.level);
}
void main() {
  print(Rank(2).compareTo(Rank(9)));
}"#,
        ["-1"]
    };

    comparable_compare_greater_returns_positive => {
        r#"class Rank implements Comparable<Rank> {
  int level;
  Rank(this.level);
  int compareTo(Rank other) => level.compareTo(other.level);
}
void main() {
  print(Rank(10).compareTo(Rank(3)));
}"#,
        ["1"]
    };

    comparable_sort_ascending_list => {
        r#"class Score implements Comparable<Score> {
  int pts;
  Score(this.pts);
  int compareTo(Score other) => pts.compareTo(other.pts);
}
void main() {
  var list = [Score(30), Score(10), Score(20)];
  list.sort();
  print(list[0].pts);
  print(list[2].pts);
}"#,
        ["10", "30"]
    };

    comparable_sort_descending_via_comparator => {
        r#"class Score implements Comparable<Score> {
  int pts;
  Score(this.pts);
  int compareTo(Score other) => pts.compareTo(other.pts);
}
void main() {
  var list = [Score(1), Score(3), Score(2)];
  list.sort((a, b) => b.compareTo(a));
  print(list[0].pts);
}"#,
        ["3"]
    };

    comparable_less_operator_via_compare_to => {
        r#"class Age implements Comparable<Age> {
  int years;
  Age(this.years);
  int compareTo(Age other) => years.compareTo(other.years);
  bool operator <(Age other) => compareTo(other) < 0;
}
void main() {
  print(Age(10) < Age(20));
}"#,
        ["true"]
    };

    comparable_greater_operator_via_compare_to => {
        r#"class Age implements Comparable<Age> {
  int years;
  Age(this.years);
  int compareTo(Age other) => years.compareTo(other.years);
  bool operator >(Age other) => compareTo(other) > 0;
}
void main() {
  print(Age(30) > Age(5));
}"#,
        ["true"]
    };

    comparable_less_or_equal_via_compare_to => {
        r#"class Size implements Comparable<Size> {
  int n;
  Size(this.n);
  int compareTo(Size other) => n.compareTo(other.n);
  bool operator <=(Size other) => compareTo(other) <= 0;
}
void main() {
  print(Size(4) <= Size(4));
}"#,
        ["true"]
    };

    comparable_greater_or_equal_via_compare_to => {
        r#"class Size implements Comparable<Size> {
  int n;
  Size(this.n);
  int compareTo(Size other) => n.compareTo(other.n);
  bool operator >=(Size other) => compareTo(other) >= 0;
}
void main() {
  print(Size(7) >= Size(2));
}"#,
        ["true"]
    };

    comparable_equal_via_compare_to_zero => {
        r#"class Tag implements Comparable<Tag> {
  String label;
  Tag(this.label);
  int compareTo(Tag other) => label.compareTo(other.label);
}
void main() {
  print(Tag('a').compareTo(Tag('a')) == 0);
}"#,
        ["true"]
    };

    comparable_string_field_lexicographic => {
        r#"class Word implements Comparable<Word> {
  String text;
  Word(this.text);
  int compareTo(Word other) => text.compareTo(other.text);
}
void main() {
  print(Word('apple').compareTo(Word('banana')));
}"#,
        ["-1"]
    };

    comparable_sort_strings_alphabetically => {
        r#"class Word implements Comparable<Word> {
  String text;
  Word(this.text);
  int compareTo(Word other) => text.compareTo(other.text);
}
void main() {
  var words = [Word('cherry'), Word('apple'), Word('banana')];
  words.sort();
  print(words[0].text);
  print(words[2].text);
}"#,
        ["apple", "cherry"]
    };

    comparable_negative_values_order => {
        r#"class Temp implements Comparable<Temp> {
  int celsius;
  Temp(this.celsius);
  int compareTo(Temp other) => celsius.compareTo(other.celsius);
}
void main() {
  print(Temp(-5).compareTo(Temp(0)));
}"#,
        ["-1"]
    };

    comparable_zero_compare_to_zero => {
        r#"class Zero implements Comparable<Zero> {
  int v;
  Zero(this.v);
  int compareTo(Zero other) => v.compareTo(other.v);
}
void main() {
  print(Zero(0).compareTo(Zero(0)));
}"#,
        ["0"]
    };

    comparable_chained_compare_in_sort => {
        r#"class Item implements Comparable<Item> {
  int pri;
  int seq;
  Item(this.pri, this.seq);
  int compareTo(Item other) {
    var c = pri.compareTo(other.pri);
    return c != 0 ? c : seq.compareTo(other.seq);
  }
}
void main() {
  var items = [Item(2, 1), Item(1, 2), Item(1, 1)];
  items.sort();
  print(items[0].pri);
  print(items[0].seq);
  print(items[2].pri);
}"#,
        ["1", "1", "2"]
    };

    comparable_find_min_via_compare_to => {
        r#"class Val implements Comparable<Val> {
  int n;
  Val(this.n);
  int compareTo(Val other) => n.compareTo(other.n);
}
Val minOf(Val a, Val b) {
  return a.compareTo(b) <= 0 ? a : b;
}
void main() {
  print(minOf(Val(3), Val(7)).n);
}"#,
        ["3"]
    };

    comparable_find_max_via_compare_to => {
        r#"class Val implements Comparable<Val> {
  int n;
  Val(this.n);
  int compareTo(Val other) => n.compareTo(other.n);
}
Val maxOf(Val a, Val b) {
  return a.compareTo(b) >= 0 ? a : b;
}
void main() {
  print(maxOf(Val(3), Val(7)).n);
}"#,
        ["7"]
    };

    comparable_list_is_sorted_check => {
        r#"class Step implements Comparable<Step> {
  int s;
  Step(this.s);
  int compareTo(Step other) => s.compareTo(other.s);
}
bool isSorted(List<Step> list) {
  for (var i = 1; i < list.length; i++) {
    if (list[i - 1].compareTo(list[i]) > 0) return false;
  }
  return true;
}
void main() {
  print(isSorted([Step(1), Step(2), Step(3)]));
}"#,
        ["true"]
    };

    comparable_list_not_sorted_detected => {
        r#"class Step implements Comparable<Step> {
  int s;
  Step(this.s);
  int compareTo(Step other) => s.compareTo(other.s);
}
bool isSorted(List<Step> list) {
  for (var i = 1; i < list.length; i++) {
    if (list[i - 1].compareTo(list[i]) > 0) return false;
  }
  return true;
}
void main() {
  print(isSorted([Step(3), Step(1)]));
}"#,
        ["false"]
    };

    comparable_operator_less_false_when_equal => {
        r#"class Box implements Comparable<Box> {
  int v;
  Box(this.v);
  int compareTo(Box other) => v.compareTo(other.v);
  bool operator <(Box other) => compareTo(other) < 0;
}
void main() {
  print(Box(5) < Box(5));
}"#,
        ["false"]
    };

    comparable_operator_greater_false_when_equal => {
        r#"class Box implements Comparable<Box> {
  int v;
  Box(this.v);
  int compareTo(Box other) => v.compareTo(other.v);
  bool operator >(Box other) => compareTo(other) > 0;
}
void main() {
  print(Box(5) > Box(5));
}"#,
        ["false"]
    };

    comparable_sort_single_element => {
        r#"class Solo implements Comparable<Solo> {
  int x;
  Solo(this.x);
  int compareTo(Solo other) => x.compareTo(other.x);
}
void main() {
  var list = [Solo(42)];
  list.sort();
  print(list[0].x);
}"#,
        ["42"]
    };

    comparable_sort_two_elements_swap => {
        r#"class Duo implements Comparable<Duo> {
  int d;
  Duo(this.d);
  int compareTo(Duo other) => d.compareTo(other.d);
}
void main() {
  var list = [Duo(2), Duo(1)];
  list.sort();
  print(list[0].d);
  print(list[1].d);
}"#,
        ["1", "2"]
    };

    comparable_reverse_sort_three_items => {
        r#"class Num implements Comparable<Num> {
  int n;
  Num(this.n);
  int compareTo(Num other) => n.compareTo(other.n);
}
void main() {
  var list = [Num(1), Num(2), Num(3)];
  list.sort((a, b) => b.compareTo(a));
  print(list.map((e) => e.n).join(','));
}"#,
        ["3,2,1"]
    };

    comparable_compare_with_large_values => {
        r#"class Big implements Comparable<Big> {
  int v;
  Big(this.v);
  int compareTo(Big other) => v.compareTo(other.v);
}
void main() {
  print(Big(1000000).compareTo(Big(999999)));
}"#,
        ["1"]
    };

    comparable_string_sort_case_sensitive => {
        r#"class Name implements Comparable<Name> {
  String s;
  Name(this.s);
  int compareTo(Name other) => s.compareTo(other.s);
}
void main() {
  print(Name('B').compareTo(Name('a')));
}"#,
        ["-1"]
    };

    comparable_median_of_three => {
        r#"class Mid implements Comparable<Mid> {
  int m;
  Mid(this.m);
  int compareTo(Mid other) => m.compareTo(other.m);
}
Mid median(Mid a, Mid b, Mid c) {
  var list = [a, b, c];
  list.sort();
  return list[1];
}
void main() {
  print(median(Mid(3), Mid(1), Mid(2)).m);
}"#,
        ["2"]
    };

    comparable_less_equal_mixed_operators => {
        r#"class Point implements Comparable<Point> {
  int x;
  Point(this.x);
  int compareTo(Point other) => x.compareTo(other.x);
  bool operator <(Point o) => compareTo(o) < 0;
  bool operator <=(Point o) => compareTo(o) <= 0;
}
void main() {
  print(Point(1) < Point(2));
  print(Point(2) <= Point(2));
}"#,
        ["true", "true"]
    };

    comparable_greater_equal_mixed_operators => {
        r#"class Point implements Comparable<Point> {
  int x;
  Point(this.x);
  int compareTo(Point other) => x.compareTo(other.x);
  bool operator >(Point o) => compareTo(o) > 0;
  bool operator >=(Point o) => compareTo(o) >= 0;
}
void main() {
  print(Point(5) > Point(1));
  print(Point(3) >= Point(3));
}"#,
        ["true", "true"]
    };

    comparable_sort_preserves_equal_elements => {
        r#"class Key implements Comparable<Key> {
  int k;
  Key(this.k);
  int compareTo(Key other) => k.compareTo(other.k);
}
void main() {
  var list = [Key(2), Key(1), Key(2)];
  list.sort();
  print(list[0].k);
  print(list[1].k);
  print(list[2].k);
}"#,
        ["1", "2", "2"]
    };

    comparable_compare_reflexive => {
        r#"class Ref implements Comparable<Ref> {
  int r;
  Ref(this.r);
  int compareTo(Ref other) => r.compareTo(other.r);
}
void main() {
  var x = Ref(7);
  print(x.compareTo(x));
}"#,
        ["0"]
    };

    comparable_compare_transitive_less => {
        r#"class T implements Comparable<T> {
  int t;
  T(this.t);
  int compareTo(T other) => t.compareTo(other.t);
}
void main() {
  var a = T(1);
  var b = T(2);
  var c = T(3);
  print(a.compareTo(b) < 0 && b.compareTo(c) < 0);
}"#,
        ["true"]
    };

    comparable_sort_empty_list => {
        r#"class E implements Comparable<E> {
  int e;
  E(this.e);
  int compareTo(E other) => e.compareTo(other.e);
}
void main() {
  var list = <E>[];
  list.sort();
  print(list.length);
}"#,
        ["0"]
    };

    comparable_operator_not_equal_via_compare => {
        r#"class Id implements Comparable<Id> {
  int id;
  Id(this.id);
  int compareTo(Id other) => id.compareTo(other.id);
}
void main() {
  print(Id(1).compareTo(Id(2)) != 0);
}"#,
        ["true"]
    };

    comparable_priority_queue_style_insert => {
        r#"class Task implements Comparable<Task> {
  int priority;
  Task(this.priority);
  int compareTo(Task other) => priority.compareTo(other.priority);
}
void main() {
  var tasks = [Task(3), Task(1), Task(2)];
  tasks.sort();
  print(tasks.first.priority);
}"#,
        ["1"]
    };

    comparable_version_tuple_style => {
        r#"class Ver implements Comparable<Ver> {
  int major;
  int minor;
  Ver(this.major, this.minor);
  int compareTo(Ver other) {
    var c = major.compareTo(other.major);
    return c != 0 ? c : minor.compareTo(other.minor);
  }
}
void main() {
  print(Ver(2, 0).compareTo(Ver(1, 9)));
}"#,
        ["1"]
    };

    comparable_version_sort_order => {
        r#"class Ver implements Comparable<Ver> {
  int major;
  int minor;
  Ver(this.major, this.minor);
  int compareTo(Ver other) {
    var c = major.compareTo(other.major);
    return c != 0 ? c : minor.compareTo(other.minor);
  }
}
void main() {
  var vers = [Ver(2, 1), Ver(1, 10), Ver(2, 0)];
  vers.sort();
  print(vers[0].major);
  print(vers[0].minor);
  print(vers[2].major);
}"#,
        ["1", "10", "2"]
    };

    comparable_less_operator_false_when_greater => {
        r#"class W implements Comparable<W> {
  int w;
  W(this.w);
  int compareTo(W other) => w.compareTo(other.w);
  bool operator <(W other) => compareTo(other) < 0;
}
void main() {
  print(W(9) < W(1));
}"#,
        ["false"]
    };

    comparable_greater_operator_false_when_less => {
        r#"class W implements Comparable<W> {
  int w;
  W(this.w);
  int compareTo(W other) => w.compareTo(other.w);
  bool operator >(W other) => compareTo(other) > 0;
}
void main() {
  print(W(1) > W(9));
}"#,
        ["false"]
    };

    comparable_binary_search_style => {
        r#"class Slot implements Comparable<Slot> {
  int s;
  Slot(this.s);
  int compareTo(Slot other) => s.compareTo(other.s);
}
int indexOf(List<Slot> list, Slot target) {
  for (var i = 0; i < list.length; i++) {
    if (list[i].compareTo(target) == 0) return i;
  }
  return -1;
}
void main() {
  var list = [Slot(10), Slot(20), Slot(30)];
  print(indexOf(list, Slot(20)));
}"#,
        ["1"]
    };

    comparable_sort_then_first_is_min => {
        r#"class Min implements Comparable<Min> {
  int v;
  Min(this.v);
  int compareTo(Min other) => v.compareTo(other.v);
}
void main() {
  var list = [Min(5), Min(-1), Min(3)];
  list.sort();
  print(list.first.v);
}"#,
        ["-1"]
    };

    comparable_sort_then_last_is_max => {
        r#"class Max implements Comparable<Max> {
  int v;
  Max(this.v);
  int compareTo(Max other) => v.compareTo(other.v);
}
void main() {
  var list = [Max(5), Max(-1), Max(3)];
  list.sort();
  print(list.last.v);
}"#,
        ["5"]
    };

    comparable_compare_negative_one_less => {
        r#"class N implements Comparable<N> {
  int n;
  N(this.n);
  int compareTo(N other) => n.compareTo(other.n);
}
void main() {
  print(N(-10).compareTo(N(-5)) < 0);
}"#,
        ["true"]
    };

    comparable_compare_positive_one_greater => {
        r#"class N implements Comparable<N> {
  int n;
  N(this.n);
  int compareTo(N other) => n.compareTo(other.n);
}
void main() {
  print(N(10).compareTo(N(5)) > 0);
}"#,
        ["true"]
    };

    comparable_all_operators_consistent => {
        r#"class Ord implements Comparable<Ord> {
  int o;
  Ord(this.o);
  int compareTo(Ord other) => o.compareTo(other.o);
  bool operator <(Ord x) => compareTo(x) < 0;
  bool operator <=(Ord x) => compareTo(x) <= 0;
  bool operator >(Ord x) => compareTo(x) > 0;
  bool operator >=(Ord x) => compareTo(x) >= 0;
}
void main() {
  var a = Ord(2);
  var b = Ord(5);
  print(a < b);
  print(a <= b);
  print(b > a);
  print(b >= a);
}"#,
        ["true", "true", "true", "true"]
    };
}
