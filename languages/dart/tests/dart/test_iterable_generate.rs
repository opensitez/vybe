//! Iterable.generate and List.generate factories, growable vs fixed-length lists,
//! and lazy iteration patterns over generated sequences.

dart_cases! {
    iterable_generate_basic_indices => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) => i);
  print(seq.join(','));
}"#,
        ["0,1,2,3"]
    };

    iterable_generate_zero_length_is_empty => {
        r#"void main() {
  var seq = Iterable.generate(0, (i) => i);
  print(seq.isEmpty);
}"#,
        ["true"]
    };

    iterable_generate_length_one => {
        r#"void main() {
  var seq = Iterable.generate(1, (i) => i + 10);
  print(seq.first);
}"#,
        ["10"]
    };

    iterable_generate_squares => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i * i);
  print(seq.join(','));
}"#,
        ["0,1,4,9,16"]
    };

    iterable_generate_constant_value => {
        r#"void main() {
  var seq = Iterable.generate(3, (_) => 7);
  print(seq.join(','));
}"#,
        ["7,7,7"]
    };

    iterable_generate_doubles => {
        r#"void main() {
  var seq = Iterable.generate(3, (i) => i * 0.5);
  print(seq.join(','));
}"#,
        ["0.0,0.5,1.0"]
    };

    iterable_generate_strings_from_index => {
        r#"void main() {
  var seq = Iterable.generate(3, (i) => 'n$i');
  print(seq.join('-'));
}"#,
        ["n0-n1-n2"]
    };

    iterable_generate_take_first_three => {
        r#"void main() {
  var seq = Iterable.generate(10, (i) => i);
  print(seq.take(3).join(','));
}"#,
        ["0,1,2"]
    };

    iterable_generate_skip_prefix => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i);
  print(seq.skip(2).join(','));
}"#,
        ["2,3,4"]
    };

    iterable_generate_map_doubles => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) => i);
  print(seq.map((n) => n * 2).join(','));
}"#,
        ["0,2,4,6"]
    };

    iterable_generate_where_filter => {
        r#"void main() {
  var seq = Iterable.generate(6, (i) => i);
  print(seq.where((n) => n % 2 == 0).join(','));
}"#,
        ["0,2,4"]
    };

    iterable_generate_fold_sum => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i + 1);
  print(seq.fold(0, (acc, n) => acc + n));
}"#,
        ["15"]
    };

    iterable_generate_reduce_product => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) => i + 1);
  print(seq.reduce((a, b) => a * b));
}"#,
        ["24"]
    };

    iterable_generate_to_list_materialize => {
        r#"void main() {
  var list = Iterable.generate(3, (i) => i * 3).toList();
  print(list.join(','));
}"#,
        ["0,3,6"]
    };

    iterable_generate_element_at_index => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i);
  print(seq.elementAt(3));
}"#,
        ["3"]
    };

    iterable_generate_powers_of_two => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) {
    var p = 1;
    for (var j = 0; j < i; j++) {
      p = p * 2;
    }
    return p;
  });
  print(seq.join(','));
}"#,
        ["1,2,4,8"]
    };

    iterable_generate_triangle_numbers => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) => (i * (i + 1)) ~/ 2);
  print(seq.join(','));
}"#,
        ["0,1,3,6"]
    };

    iterable_generate_alternating_signs => {
        r#"void main() {
  var seq = Iterable.generate(4, (i) => i % 2 == 0 ? 1 : -1);
  print(seq.join(','));
}"#,
        ["1,-1,1,-1"]
    };

    iterable_generate_modulo_pattern => {
        r#"void main() {
  var seq = Iterable.generate(6, (i) => i % 3);
  print(seq.join(','));
}"#,
        ["0,1,2,0,1,2"]
    };

    iterable_generate_lazy_not_all_consumed => {
        r#"void main() {
  var seq = Iterable.generate(100, (i) => i);
  print(seq.take(2).join(','));
}"#,
        ["0,1"]
    };

    iterable_generate_for_in_loop => {
        r#"void main() {
  var sum = 0;
  for (var n in Iterable.generate(4, (i) => i + 1)) {
    sum = sum + n;
  }
  print(sum);
}"#,
        ["10"]
    };

    iterable_generate_contains_check => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i * 2);
  print(seq.contains(6));
}"#,
        ["true"]
    };

    iterable_generate_every_predicate => {
        r#"void main() {
  var seq = Iterable.generate(3, (i) => i + 1);
  print(seq.every((n) => n > 0));
}"#,
        ["true"]
    };

    iterable_generate_any_predicate => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i);
  print(seq.any((n) => n == 4));
}"#,
        ["true"]
    };

    list_generate_basic_indices => {
        r#"void main() {
  var list = List.generate(4, (i) => i);
  print(list.join(','));
}"#,
        ["0,1,2,3"]
    };

    list_generate_growable_default_length => {
        r#"void main() {
  var list = List.generate(3, (i) => i);
  print(list.length);
}"#,
        ["3"]
    };

    list_generate_growable_true_add_element => {
        r#"void main() {
  var list = List.generate(2, (i) => i, growable: true);
  list.add(99);
  print(list.join(','));
}"#,
        ["0,1,99"]
    };

    list_generate_growable_false_fixed_length => {
        r#"void main() {
  var list = List.generate(3, (i) => i, growable: false);
  print(list.length);
}"#,
        ["3"]
    };

    list_generate_zero_length => {
        r#"void main() {
  var list = List.generate(0, (i) => i);
  print(list.length);
  print(list.isEmpty);
}"#,
        ["0", "true"]
    };

    list_generate_length_one => {
        r#"void main() {
  var list = List.generate(1, (i) => 42);
  print(list[0]);
}"#,
        ["42"]
    };

    list_generate_squares => {
        r#"void main() {
  var list = List.generate(5, (i) => i * i);
  print(list.join(','));
}"#,
        ["0,1,4,9,16"]
    };

    list_generate_identity_function => {
        r#"void main() {
  var list = List.generate(4, (i) => i);
  print(list[2]);
}"#,
        ["2"]
    };

    list_generate_strings => {
        r#"void main() {
  var list = List.generate(3, (i) => 'x');
  print(list.join(''));
}"#,
        ["xxx"]
    };

    list_generate_with_index_offset => {
        r#"void main() {
  var list = List.generate(3, (i) => i + 10);
  print(list.join(','));
}"#,
        ["10,11,12"]
    };

    list_generate_factorial_row => {
        r#"void main() {
  var list = List.generate(4, (i) {
    var f = 1;
    for (var j = 1; j <= i; j++) {
      f = f * j;
    }
    return f;
  });
  print(list.join(','));
}"#,
        ["1,1,2,6"]
    };

    list_generate_fixed_growable_false_index_access => {
        r#"void main() {
  var list = List.generate(4, (i) => i * 2, growable: false);
  print(list[3]);
}"#,
        ["6"]
    };

    list_generate_nested_list_structure => {
        r#"void main() {
  var list = List.generate(2, (i) => List.generate(2, (j) => i + j));
  print(list[1][1]);
}"#,
        ["2"]
    };

    list_generate_doubles => {
        r#"void main() {
  var list = List.generate(3, (i) => i * 1.5);
  print(list.join(','));
}"#,
        ["0.0,1.5,3.0"]
    };

    list_generate_chars_from_codes => {
        r#"void main() {
  var list = List.generate(3, (i) => String.fromCharCode(97 + i));
  print(list.join(''));
}"#,
        ["abc"]
    };

    list_generate_constant_fill => {
        r#"void main() {
  var list = List.generate(5, (_) => 'z');
  print(list.length);
  print(list[4]);
}"#,
        ["5", "z"]
    };

    compare_iterable_and_list_generate_same_values => {
        r#"void main() {
  var lazy = Iterable.generate(4, (i) => i * 2);
  var eager = List.generate(4, (i) => i * 2);
  print(lazy.join(','));
  print(eager.join(','));
}"#,
        ["0,2,4,6", "0,2,4,6"]
    };

    list_generate_growable_vs_fixed_same_contents => {
        r#"void main() {
  var grow = List.generate(3, (i) => i, growable: true);
  var fixed = List.generate(3, (i) => i, growable: false);
  print(grow.join(','));
  print(fixed.join(','));
}"#,
        ["0,1,2", "0,1,2"]
    };

    iterable_generate_skip_while_tail => {
        r#"void main() {
  var seq = Iterable.generate(6, (i) => i);
  print(seq.skipWhile((n) => n < 3).join(','));
}"#,
        ["3,4,5"]
    };

    iterable_generate_take_while_prefix => {
        r#"void main() {
  var seq = Iterable.generate(6, (i) => i);
  print(seq.takeWhile((n) => n < 4).join(','));
}"#,
        ["0,1,2,3"]
    };

    iterable_generate_followed_by_suffix => {
        r#"void main() {
  var seq = Iterable.generate(2, (i) => i).followedBy([9, 10]);
  print(seq.join(','));
}"#,
        ["0,1,9,10"]
    };

    list_generate_map_then_join => {
        r#"void main() {
  var list = List.generate(4, (i) => i);
  print(list.map((n) => n + 1).join(','));
}"#,
        ["1,2,3,4"]
    };

    list_generate_where_then_length => {
        r#"void main() {
  var list = List.generate(6, (i) => i);
  print(list.where((n) => n % 2 == 1).length);
}"#,
        ["3"]
    };

    iterable_generate_expand_nested => {
        r#"void main() {
  var seq = Iterable.generate(2, (i) => [i, i + 1]);
  print(seq.expand((pair) => pair).join(','));
}"#,
        ["0,1,1,2"]
    };

    list_generate_reversed_join => {
        r#"void main() {
  var list = List.generate(3, (i) => i + 1);
  print(list.reversed.join(','));
}"#,
        ["3,2,1"]
    };

    iterable_generate_first_and_last => {
        r#"void main() {
  var seq = Iterable.generate(5, (i) => i + 1);
  print(seq.first);
  print(seq.last);
}"#,
        ["1", "5"]
    };

    list_generate_sublist_slice => {
        r#"void main() {
  var list = List.generate(5, (i) => i);
  print(list.sublist(1, 4).join(','));
}"#,
        ["1,2,3"]
    };
}
