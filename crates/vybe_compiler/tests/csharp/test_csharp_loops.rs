//! All loop forms and jump semantics in C#.
use super::helpers::run_csharp;

#[test]
fn for_loop_counts_with_pre_increment() {
    assert_eq!(
        run_csharp(r#"int s=0; for(int i=1;i<=4;i++) s+=i; Console.WriteLine(s);"#),
        &["10"]
    );
}

#[test]
fn for_loop_counts_down_with_decrement() {
    assert_eq!(
        run_csharp(r#"string r=""; for(int i=3;i>=1;i--) r+=i; Console.WriteLine(r);"#),
        &["321"]
    );
}

#[test]
fn for_loop_with_empty_init_and_update_acts_as_while() {
    assert_eq!(
        run_csharp(r#"int i=0; for(;i<3;) i++; Console.WriteLine(i);"#),
        &["3"]
    );
}

#[test]
fn while_loop_body_skipped_when_condition_initially_false() {
    assert_eq!(
        run_csharp(r#"int count=0; while(false) count++; Console.WriteLine(count);"#),
        &["0"]
    );
}

#[test]
fn do_while_body_runs_at_least_once_when_condition_false() {
    assert_eq!(
        run_csharp(r#"int count=0; do { count++; } while(false); Console.WriteLine(count);"#),
        &["1"]
    );
}

#[test]
fn foreach_iterates_array_in_declaration_order() {
    assert_eq!(
        run_csharp(r#"int s=0; foreach(var x in new[]{3,1,4}) s+=x; Console.WriteLine(s);"#),
        &["8"]
    );
}

#[test]
fn foreach_over_string_visits_each_char() {
    assert_eq!(
        run_csharp(r#"int n=0; foreach(char c in "hello") n++; Console.WriteLine(n);"#),
        &["5"]
    );
}

#[test]
fn break_exits_innermost_loop_only() {
    assert_eq!(
        run_csharp(
            r#"
int total = 0;
for(int i=0;i<3;i++) {
    for(int j=0;j<3;j++) {
        if(j==1) break;
        total++;
    }
}
Console.WriteLine(total);
"#
        ),
        &["3"]
    );
}

#[test]
fn continue_skips_rest_of_body_and_re_evaluates_condition() {
    assert_eq!(
        run_csharp(r#"int s=0; for(int i=0;i<5;i++) { if(i%2==0) continue; s+=i; } Console.WriteLine(s);"#),
        &["4"]
    );
}

#[test]
fn goto_jumps_forward_to_labeled_statement() {
    assert_eq!(
        run_csharp(
            r#"
int x = 0;
goto done;
x = 99;
done:
Console.WriteLine(x);
"#
        ),
        &["0"]
    );
}

#[test]
fn nested_foreach_produces_cartesian_pair_count() {
    assert_eq!(
        run_csharp(
            r#"
int count=0;
foreach(var a in new[]{1,2})
    foreach(var b in new[]{1,2,3})
        count++;
Console.WriteLine(count);
"#
        ),
        &["6"]
    );
}

#[test]
fn while_accumulates_with_compound_assignment() {
    assert_eq!(
        run_csharp(r#"int n=1; while(n<100) n*=2; Console.WriteLine(n);"#),
        &["128"]
    );
}
