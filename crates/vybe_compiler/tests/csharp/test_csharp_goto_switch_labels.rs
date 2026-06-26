//! `goto case`, `goto default`, named labels, and `break`/`continue` in loops.

csharp_cases! {
    goto_case_falls_through_from_first_to_second => {
        r#"int code = 1;
string trace = "";
switch (code) {
    case 1:
        trace += "A";
        goto case 2;
    case 2:
        trace += "B";
        break;
}
Console.WriteLine(trace);"#,
        ["AB"]
    };

    goto_case_chains_three_cases => {
        r#"int n = 1;
string r = "";
switch (n) {
    case 1: r += "1"; goto case 2;
    case 2: r += "2"; goto case 3;
    case 3: r += "3"; break;
}
Console.WriteLine(r);"#,
        ["123"]
    };

    goto_case_from_two_to_four_skipping_three => {
        r#"int n = 2;
string r = "";
switch (n) {
    case 1: r += "1"; goto case 4;
    case 2: r += "2"; goto case 4;
    case 3: r += "3"; break;
    case 4: r += "4"; break;
}
Console.WriteLine(r);"#,
        ["24"]
    };

    goto_default_from_non_matching_case_value => {
        r#"int n = 99;
string label = "";
switch (n) {
    case 1: label = "one"; break;
    case 2: label = "two"; break;
    default:
        label = "other";
        break;
}
Console.WriteLine(label);"#,
        ["other"]
    };

    goto_default_from_case_arm => {
        r#"int n = 1;
string r = "";
switch (n) {
    case 1:
        r += "start";
        goto default;
    default:
        r += ":default";
        break;
}
Console.WriteLine(r);"#,
        ["start:default"]
    };

    goto_default_then_break_exits_switch => {
        r#"int n = 0;
string r = "";
switch (n) {
    case 0:
        goto default;
    default:
        r = "done";
        break;
}
Console.WriteLine(r);"#,
        ["done"]
    };

    goto_label_jumps_forward_over_code => {
        r#"int x = 0;
start:
x++;
if (x < 3) goto start;
Console.WriteLine(x);"#,
        ["3"]
    };

    goto_label_jumps_to_shared_cleanup => {
        r#"int n = 1;
string msg = "";
if (n == 1) goto cleanup;
msg = "skip";
cleanup:
msg = "ok";
Console.WriteLine(msg);"#,
        ["ok"]
    };

    break_exits_inner_for_only => {
        r#"int total = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) break;
        total++;
    }
}
Console.WriteLine(total);"#,
        ["3"]
    };

    continue_skips_odd_additions_in_for => {
        r#"int sum = 0;
for (int i = 1; i <= 6; i++) {
    if (i % 2 == 0) continue;
    sum += i;
}
Console.WriteLine(sum);"#,
        ["9"]
    };

    break_in_while_exits_loop => {
        r#"int n = 0;
while (true) {
    n++;
    if (n == 3) break;
}
Console.WriteLine(n);"#,
        ["3"]
    };

    continue_in_while_skips_iteration => {
        r#"int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n == 3) continue;
    sum += n;
}
Console.WriteLine(sum);"#,
        ["8"]
    };

    break_in_do_while_exits => {
        r#"int n = 0;
do {
    n++;
    if (n == 2) break;
} while (n < 10);
Console.WriteLine(n);"#,
        ["2"]
    };

    continue_in_do_while_skips_body => {
        r#"int n = 0;
int sum = 0;
do {
    n++;
    if (n == 2) continue;
    sum += n;
} while (n < 4);
Console.WriteLine(sum);"#,
        ["7"]
    };

    labeled_continue_in_for_loop => {
        r#"int sum = 0;
for (int i = 0; i < 5; i++) {
    if (i == 2) continue;
    sum += i;
}
Console.WriteLine(sum);"#,
        ["8"]
    };

    goto_label_exits_nested_loop_early => {
        r#"int count = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 3; j++) {
        if (i == 1 && j == 1) goto finished;
        count++;
    }
}
finished:
Console.WriteLine(count);"#,
        ["7"]
    };

    break_in_switch_inside_loop_runs_once => {
        r#"string report = "";
for (int i = 0; i < 3; i++) {
    switch (i) {
        case 0: report += "a"; break;
        case 1: report += "b"; break;
        case 2: report += "c"; break;
    }
}
Console.WriteLine(report);"#,
        ["abc"]
    };

    goto_case_with_string_switch => {
        r#"string key = "b";
string r = "";
switch (key) {
    case "a": r += "A"; goto case "b";
    case "b": r += "B"; break;
    case "c": r += "C"; break;
}
Console.WriteLine(r);"#,
        ["B"]
    };

    goto_default_on_string_switch => {
        r#"string key = "z";
string r = "";
switch (key) {
    case "a": r = "A"; break;
    case "b": r = "B"; break;
    default: r = "?"; break;
}
Console.WriteLine(r);"#,
        ["?"]
    };

    nested_loop_break_does_not_exit_outer => {
        r#"int outerRuns = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        if (j == 1) break;
        outerRuns++;
    }
}
Console.WriteLine(outerRuns);"#,
        ["2"]
    };

    nested_loop_continue_affects_inner_only => {
        r#"int hits = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) continue;
        hits++;
    }
}
Console.WriteLine(hits);"#,
        ["4"]
    };

    goto_label_before_switch_merge_point => {
        r#"int mode = 1;
string result = "";
if (mode == 0) goto merge;
switch (mode) {
    case 1: result += "one"; break;
    case 2: result += "two"; break;
}
merge:
result += "!";
Console.WriteLine(result);"#,
        ["one!"]
    };

    switch_fallthrough_via_goto_case_only => {
        r#"int v = 1;
int total = 0;
switch (v) {
    case 1: total += 10; goto case 2;
    case 2: total += 1; break;
}
Console.WriteLine(total);"#,
        ["11"]
    };

    break_in_inner_switch_inside_loop => {
        r#"string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0: log += "in"; break;
            }
            log += ";";
            break;
        case 1: log += "out"; break;
    }
}
Console.WriteLine(log);"#,
        ["in;out"]
    };

    continue_in_foreach_skips_element => {
        r#"int sum = 0;
foreach (var x in new int[] { 1, 2, 3, 4 }) {
    if (x == 2) continue;
    sum += x;
}
Console.WriteLine(sum);"#,
        ["8"]
    };

    goto_label_shared_by_two_paths => {
        r#"int flag = 1;
int result = 0;
if (flag == 0) goto finish;
result = 5;
finish:
Console.WriteLine(result);"#,
        ["5"]
    };

    goto_case_on_enum_switch => {
        r#"enum Color { Red, Green, Blue }
Color c = Color.Red;
string name = "";
switch (c) {
    case Color.Red: name += "R"; goto case Color.Green;
    case Color.Green: name += "G"; break;
    case Color.Blue: name += "B"; break;
}
Console.WriteLine(name);"#,
        ["RG"]
    };

    goto_default_on_enum_switch => {
        r#"enum Color { Red, Green }
Color c = (Color)9;
string name = "";
switch (c) {
    case Color.Red: name = "R"; break;
    case Color.Green: name = "G"; break;
    default: name = "?"; break;
}
Console.WriteLine(name);"#,
        ["?"]
    };

    break_in_switch_does_not_exit_enclosing_loop => {
        r#"int i = 0;
while (i < 2) {
    switch (i) {
        case 0: i++; break;
        case 1: i++; break;
    }
}
Console.WriteLine(i);"#,
        ["2"]
    };

    continue_in_for_with_labeled_logic => {
        r#"string chars = "";
for (int i = 0; i < 4; i++) {
    if (i == 2) continue;
    chars += i.ToString();
}
Console.WriteLine(chars);"#,
        ["013"]
    };

    goto_label_after_switch_accumulates => {
        r#"int code = 2;
int acc = 0;
switch (code) {
    case 1: acc += 1; break;
    case 2: acc += 2; goto default;
    default: acc += 100; break;
}
Console.WriteLine(acc);"#,
        ["102"]
    };

    nested_goto_case_within_same_switch => {
        r#"int x = 1;
string s = "";
switch (x) {
    case 1: s += "a"; goto case 2;
    case 2: s += "b"; goto case 3;
    case 3: s += "c"; break;
}
Console.WriteLine(s);"#,
        ["abc"]
    };

    break_vs_continue_in_same_loop => {
        r#"string r = "";
for (int i = 0; i < 4; i++) {
    if (i == 1) continue;
    if (i == 3) break;
    r += i;
}
Console.WriteLine(r);"#,
        ["02"]
    };

    goto_label_from_inner_if_in_loop => {
        r#"int sum = 0;
for (int i = 0; i < 5; i++) {
    if (i == 3) goto done;
    sum += i;
}
done:
Console.WriteLine(sum);"#,
        ["3"]
    };

    switch_goto_default_from_middle_case => {
        r#"int n = 2;
string r = "";
switch (n) {
    case 1: r += "1"; break;
    case 2: r += "2"; goto default;
    default: r += "D"; break;
}
Console.WriteLine(r);"#,
        ["2D"]
    };

    foreach_break_exits_after_first_match => {
        r#"string seen = "";
foreach (var ch in "abc") {
    seen += ch;
    if (ch == 'b') break;
}
Console.WriteLine(seen);"#,
        ["ab"]
    };

    while_continue_restarts_without_increment_bug => {
        r#"int i = 0;
int sum = 0;
while (i < 4) {
    i++;
    if (i == 2) continue;
    sum += i;
}
Console.WriteLine(sum);"#,
        ["8"]
    };

    goto_case_preserves_order_with_break => {
        r#"int k = 1;
string buf = "";
switch (k) {
    case 1: buf += "1"; goto case 2;
    case 2: buf += "2"; break;
    case 3: buf += "3"; break;
}
Console.WriteLine(buf);"#,
        ["12"]
    };

    double_nested_loop_goto_label_escape => {
        r#"int ticks = 0;
for (int a = 0; a < 2; a++) {
    for (int b = 0; b < 2; b++) {
        ticks++;
        if (ticks == 3) goto done;
    }
}
done:
Console.WriteLine(ticks);"#,
        ["3"]
    };

    switch_default_without_goto_still_runs => {
        r#"int v = 5;
string tag = "";
switch (v) {
    case 1: tag = "one"; break;
    default: tag = "many"; break;
}
Console.WriteLine(tag);"#,
        ["many"]
    };

    goto_label_switch_mix_with_loop_break => {
        r#"string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            log += "0";
            break;
        case 1:
            log += "1";
            break;
    }
    if (i == 1) break;
}
Console.WriteLine(log);"#,
        ["01"]
    };

    continue_outer_not_available_only_inner => {
        r#"int count = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        if (j == 0) continue;
        count++;
    }
}
Console.WriteLine(count);"#,
        ["2"]
    };

    goto_case_on_zero_value => {
        r#"int v = 0;
string r = "";
switch (v) {
    case 0: r += "0"; goto case 1;
    case 1: r += "1"; break;
}
Console.WriteLine(r);"#,
        ["01"]
    };

    goto_label_skips_else_branch => {
        r#"int pick = 1;
string r = "";
if (pick == 0) r = "zero";
else goto show;
show:
r = "one";
Console.WriteLine(r);"#,
        ["one"]
    };

    break_in_switch_case_stops_fallthrough => {
        r#"int n = 1;
string r = "";
switch (n) {
    case 1: r += "x"; break;
    case 2: r += "y"; break;
}
Console.WriteLine(r);"#,
        ["x"]
    };
}
