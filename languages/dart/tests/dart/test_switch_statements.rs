//! Switch statements: int, string, fall-through, break, default, enum-like constants.

dart_cases! {
    int_switch_matches_first_case => {
        r#"void main() {
  var code = 1;
  switch (code) {
    case 1:
      print('alpha');
      break;
    case 2:
      print('beta');
      break;
    default:
      print('unknown');
  }
}"#,
        ["alpha"]
    };

    int_switch_matches_middle_case => {
        r#"void main() {
  var code = 2;
  switch (code) {
    case 1:
      print('one');
      break;
    case 2:
      print('two');
      break;
    case 3:
      print('three');
      break;
    default:
      print('other');
  }
}"#,
        ["two"]
    };

    int_switch_matches_last_non_default_case => {
        r#"void main() {
  var code = 3;
  switch (code) {
    case 1:
      print('one');
      break;
    case 2:
      print('two');
      break;
    case 3:
      print('three');
      break;
    default:
      print('many');
  }
}"#,
        ["three"]
    };

    int_switch_no_match_routes_to_default => {
        r#"void main() {
  var code = 99;
  switch (code) {
    case 1:
      print('one');
      break;
    case 2:
      print('two');
      break;
    default:
      print('fallback');
  }
}"#,
        ["fallback"]
    };

    int_switch_default_only_with_literal_selector => {
        r#"void main() {
  switch (42) {
    default:
      print('only-default');
  }
}"#,
        ["only-default"]
    };

    int_switch_empty_case_body_falls_into_default => {
        r#"void main() {
  switch (5) {
    case 5:
    default:
      print('via-default');
  }
}"#,
        ["via-default"]
    };

    int_switch_fallthrough_two_labels_one_body => {
        r#"void main() {
  var tier = 2;
  switch (tier) {
    case 1:
    case 2:
      print('low');
      break;
    case 3:
      print('high');
      break;
    default:
      print('out');
  }
}"#,
        ["low"]
    };

    int_switch_fallthrough_three_labels_before_break => {
        r#"void main() {
  var score = 3;
  switch (score) {
    case 1:
    case 2:
    case 3:
      print('pass');
      break;
    default:
      print('fail');
  }
}"#,
        ["pass"]
    };

    int_switch_break_prevents_fallthrough_to_next_case => {
        r#"void main() {
  var n = 1;
  switch (n) {
    case 1:
      print('first');
      break;
    case 2:
      print('second');
      break;
    default:
      print('rest');
  }
}"#,
        ["first"]
    };

    int_switch_negative_value_matches_case => {
        r#"void main() {
  var delta = -1;
  switch (delta) {
    case -1:
      print('minus-one');
      break;
    case 0:
      print('zero');
      break;
    default:
      print('positive');
  }
}"#,
        ["minus-one"]
    };

    int_switch_zero_value_matches_case => {
        r#"void main() {
  var count = 0;
  switch (count) {
    case 0:
      print('empty');
      break;
    case 1:
      print('single');
      break;
    default:
      print('many');
  }
}"#,
        ["empty"]
    };

    int_switch_without_default_and_no_match_exits_silently => {
        r#"void main() {
  var x = 7;
  switch (x) {
    case 1:
      print('one');
      break;
    case 2:
      print('two');
      break;
  }
  print('done');
}"#,
        ["done"]
    };

    string_switch_matches_exact_token => {
        r#"void main() {
  var fruit = 'apple';
  switch (fruit) {
    case 'apple':
      print('red');
      break;
    case 'banana':
      print('yellow');
      break;
    default:
      print('other');
  }
}"#,
        ["red"]
    };

    string_switch_no_match_routes_to_default => {
        r#"void main() {
  var fruit = 'cherry';
  switch (fruit) {
    case 'apple':
      print('apple');
      break;
    case 'banana':
      print('banana');
      break;
    default:
      print('exotic');
  }
}"#,
        ["exotic"]
    };

    string_switch_fallthrough_grouped_weekday_labels => {
        r#"void main() {
  var day = 'Sat';
  switch (day) {
    case 'Sat':
    case 'Sun':
      print('weekend');
      break;
    case 'Mon':
      print('monday');
      break;
    default:
      print('weekday');
  }
}"#,
        ["weekend"]
    };

    string_switch_empty_string_label => {
        r#"void main() {
  var label = '';
  switch (label) {
    case '':
      print('blank');
      break;
    case 'x':
      print('letter');
      break;
    default:
      print('other');
  }
}"#,
        ["blank"]
    };

    string_switch_single_character_label => {
        r#"void main() {
  var ch = 'z';
  switch (ch) {
    case 'a':
      print('vowel-start');
      break;
    case 'z':
      print('last-letter');
      break;
    default:
      print('middle');
  }
}"#,
        ["last-letter"]
    };

    string_switch_break_stops_at_first_matching_case => {
        r#"void main() {
  var mode = 'read';
  switch (mode) {
    case 'read':
      print('reading');
      break;
    case 'write':
      print('writing');
      break;
    default:
      print('idle');
  }
}"#,
        ["reading"]
    };

    bool_switch_true_branch => {
        r#"void main() {
  var flag = true;
  switch (flag) {
    case true:
      print('yes');
      break;
    case false:
      print('no');
      break;
  }
}"#,
        ["yes"]
    };

    bool_switch_false_branch => {
        r#"void main() {
  var flag = false;
  switch (flag) {
    case true:
      print('yes');
      break;
    case false:
      print('no');
      break;
  }
}"#,
        ["no"]
    };

    enum_like_const_int_switch_red => {
        r#"void main() {
  const red = 0;
  const green = 1;
  const blue = 2;
  var color = red;
  switch (color) {
    case red:
      print('red');
      break;
    case green:
      print('green');
      break;
    case blue:
      print('blue');
      break;
    default:
      print('unknown');
  }
}"#,
        ["red"]
    };

    enum_like_const_int_switch_green => {
        r#"void main() {
  const red = 0;
  const green = 1;
  const blue = 2;
  var color = green;
  switch (color) {
    case red:
      print('red');
      break;
    case green:
      print('green');
      break;
    case blue:
      print('blue');
      break;
    default:
      print('unknown');
  }
}"#,
        ["green"]
    };

    enum_like_const_int_switch_unknown_hits_default => {
        r#"void main() {
  const red = 0;
  const green = 1;
  const blue = 2;
  var color = 9;
  switch (color) {
    case red:
      print('red');
      break;
    case green:
      print('green');
      break;
    case blue:
      print('blue');
      break;
    default:
      print('unknown');
  }
}"#,
        ["unknown"]
    };

    switch_nested_inner_string_match => {
        r#"void main() {
  var outer = 1;
  var inner = 'b';
  switch (outer) {
    case 1:
      switch (inner) {
        case 'a':
          print('inner-a');
          break;
        case 'b':
          print('inner-b');
          break;
        default:
          print('inner-other');
      }
      break;
    default:
      print('outer-other');
  }
}"#,
        ["inner-b"]
    };

    switch_inside_for_loop_selects_per_iteration => {
        r#"void main() {
  for (var i = 0; i < 3; i++) {
    switch (i) {
      case 0:
        print('zero');
        break;
      case 1:
        print('one');
        break;
      case 2:
        print('two');
        break;
    }
  }
}"#,
        ["zero", "one", "two"]
    };

    switch_case_body_runs_multiple_statements => {
        r#"void main() {
  var n = 2;
  switch (n) {
    case 2:
      print('start');
      print('end');
      break;
    default:
      print('other');
  }
}"#,
        ["start", "end"]
    };

    switch_default_runs_when_all_int_cases_miss => {
        r#"void main() {
  var port = 8080;
  switch (port) {
    case 80:
      print('http');
      break;
    case 443:
      print('https');
      break;
    default:
      print('custom');
  }
}"#,
        ["custom"]
    };

    switch_fallthrough_prints_shared_message_once => {
        r#"void main() {
  var level = 1;
  switch (level) {
    case 1:
    case 2:
      print('warn');
      break;
    case 3:
      print('error');
      break;
    default:
      print('info');
  }
}"#,
        ["warn"]
    };

    switch_variable_selector_picks_matching_int_case => {
        r#"void main() {
  var choice = 4;
  var picked = choice * 1;
  switch (picked) {
    case 4:
      print('four');
      break;
    case 8:
      print('eight');
      break;
    default:
      print('other');
  }
}"#,
        ["four"]
    };

    switch_string_variable_selects_weekday => {
        r#"void main() {
  var day = 'Wed';
  switch (day) {
    case 'Mon':
      print('start');
      break;
    case 'Wed':
      print('midweek');
      break;
    case 'Fri':
      print('end');
      break;
    default:
      print('other');
  }
}"#,
        ["midweek"]
    };

    switch_only_default_case_prints_once => {
        r#"void main() {
  var token = 'anything';
  switch (token) {
    default:
      print('catch-all');
  }
}"#,
        ["catch-all"]
    };

    switch_first_empty_case_falls_into_second_with_action => {
        r#"void main() {
  var key = 10;
  switch (key) {
    case 10:
    case 20:
      print('tens');
      break;
    default:
      print('other');
  }
}"#,
        ["tens"]
    };

    switch_break_exits_after_first_matching_fallthrough_group => {
        r#"void main() {
  var code = 3;
  switch (code) {
    case 1:
    case 2:
      print('small');
      break;
    case 3:
    case 4:
      print('medium');
      break;
    default:
      print('large');
  }
}"#,
        ["medium"]
    };

    switch_int_first_case_skips_later_unmatched_labels => {
        r#"void main() {
  var n = 1;
  switch (n) {
    case 1:
      print('hit-one');
      break;
    case 2:
      print('hit-two');
      break;
    case 3:
      print('hit-three');
      break;
  }
}"#,
        ["hit-one"]
    };

    switch_string_case_sensitive_miss_uses_default => {
        r#"void main() {
  var word = 'Hello';
  switch (word) {
    case 'hello':
      print('lower');
      break;
    case 'HELLO':
      print('upper');
      break;
    default:
      print('mixed');
  }
}"#,
        ["mixed"]
    };

    switch_selector_from_arithmetic_expression => {
        r#"void main() {
  var base = 3;
  var limit = 4;
  switch (base + limit) {
    case 7:
      print('seven');
      break;
    case 8:
      print('eight');
      break;
    default:
      print('other');
  }
}"#,
        ["seven"]
    };

    switch_default_position_does_not_affect_matching => {
        r#"void main() {
  var id = 2;
  switch (id) {
    case 1:
      print('one');
      break;
    default:
      print('default-first');
      break;
    case 2:
      print('two');
      break;
  }
}"#,
        ["two"]
    };

    switch_fallthrough_accumulates_counter_before_break => {
        r#"void main() {
  var step = 2;
  var total = 0;
  switch (step) {
    case 1:
      total = total + 1;
    case 2:
      total = total + 2;
    case 3:
      total = total + 3;
      break;
    default:
      total = total + 0;
  }
  print(total);
}"#,
        // A case WITH a body breaks implicitly — Dart 3 rejects falling out of
        // one, so `case 2` does not continue into `case 3`. Verified against
        // the Dart SDK: this program prints 2. Only an EMPTY case falls
        // through, which `switch_first_empty_case_falls_into_second_with_action`
        // covers.
        ["2"]
    };

    switch_break_inside_switch_does_not_continue_outer_loop => {
        r#"void main() {
  for (var i = 0; i < 2; i++) {
    switch (i) {
      case 0:
        print('loop-0');
        break;
      case 1:
        print('loop-1');
        break;
    }
    print('after-$i');
  }
}"#,
        ["loop-0", "after-0", "loop-1", "after-1"]
    };

    switch_empty_body_cases_all_fall_to_default_print => {
        r#"void main() {
  switch (2) {
    case 1:
    case 2:
    default:
      print('reached-default');
  }
}"#,
        ["reached-default"]
    };
}
