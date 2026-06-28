//! `goto`, labeled statements, `goto case`, and break/continue in nested loops.
use super::helpers::run_csharp;

#[test]
fn goto_jumps_to_labeled_statement() {
    assert_eq!(
        run_csharp(
            r#"int i=0;
start:
if(i<5){i++;goto start;}
Console.WriteLine(i);"#
        ),
        &["5"]
    );
}

#[test]
fn goto_in_switch_falls_through_via_goto_case() {
    assert_eq!(
        run_csharp(
            r#"int n=1;
string r="";
switch(n){
    case 1: r+="one"; goto case 2;
    case 2: r+="two"; break;
    case 3: r+="three"; break;
}
Console.WriteLine(r);"#
        ),
        &["onetwo"]
    );
}

#[test]
fn break_exits_only_innermost_loop() {
    assert_eq!(
        run_csharp(
            r#"int count=0;
for(int i=0;i<3;i++){
    for(int j=0;j<3;j++){
        if(j==1) break;
        count++;
    }
}
Console.WriteLine(count);"#
        ),
        &["3"]
    );
}

#[test]
fn continue_skips_rest_of_current_iteration() {
    assert_eq!(
        run_csharp(
            r#"int sum=0;
for(int i=1;i<=10;i++){
    if(i%2==0) continue;
    sum+=i;
}
Console.WriteLine(sum);"#
        ),
        &["25"]
    );
}
