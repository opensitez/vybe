/// Deeply nested loops and labeled-break simulation via flags/Break.
use super::helpers::run_pascal;

#[test]
fn triple_for_cell_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,k,c:Integer; begin c:=0; for i:=1 to 2 do for j:=1 to 2 do for k:=1 to 2 do Inc(c); WriteLn(c); end."#
        ),
        &["8"]
    );
}

#[test]
fn triple_for_coord_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y,z:Integer; begin for x:=1 to 2 do for y:=1 to 2 do for z:=1 to 2 do WriteLn(IntToStr(x)+'-'+IntToStr(y)+'-'+IntToStr(z)); end."#
        ),
        &["1-1-1", "1-1-2", "1-2-1", "1-2-2", "2-1-1", "2-1-2", "2-2-1", "2-2-2"]
    );
}

#[test]
fn quad_nested_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c,d,s:Integer; begin s:=0; for a:=0 to 1 do for b:=0 to 1 do for c:=0 to 1 do for d:=0 to 1 do Inc(s); WriteLn(s); end."#
        ),
        &["16"]
    );
}

#[test]
fn while_inside_for_rows() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 3 do begin j:=0; while j<i do begin WriteLn(i*10+j); Inc(j); end; end; end."#
        ),
        &["11", "12", "21", "22", "23", "31", "32", "33"]
    );
}

#[test]
fn repeat_inside_for_triangular() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,n,s:Integer; begin s:=0; for i:=1 to 4 do begin n:=0; repeat Inc(n); s:=s+n; until n=i; end; WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn for_inside_while_halving() {
    assert_eq!(
        run_pascal(
            r#"program T; var n,i:Integer; begin n:=8; while n>0 do begin for i:=1 to 2 do WriteLn(n); n:=n div 2; end; end."#
        ),
        &["8", "8", "4", "4", "2", "2", "1", "1"]
    );
}

#[test]
fn break_inner_for_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 3 do begin for j:=1 to 3 do begin if j=2 then Break; WriteLn(IntToStr(i)+','+IntToStr(j)); end; end; end."#
        ),
        &["1,1", "2,1", "3,1"]
    );
}

#[test]
fn flag_break_outer_at_target() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; done:Boolean; begin done:=false; for i:=1 to 5 do if not done then for j:=1 to 5 do if (i=2) and (j=3) then done:=true else WriteLn(IntToStr(i)+':'+IntToStr(j)); WriteLn('stop'); end."#
        ),
        &["1:1", "1:2", "1:3", "1:4", "1:5", "2:1", "2:2", "stop"]
    );
}

#[test]
fn skip_even_values_nested_style() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin for i:=1 to 6 do if i mod 2<>0 then WriteLn(i); end."#
        ),
        &["1", "3", "5"]
    );
}

#[test]
fn nested_repeat_power_of_two() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,p:Integer; begin a:=0; p:=1; repeat Inc(a); b:=0; repeat Inc(b); p:=p*2; until b=3; until a=2; WriteLn(p); end."#
        ),
        &["64"]
    );
}

#[test]
fn triple_downto_values() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,k:Integer; begin for i:=2 downto 1 do for j:=2 downto 1 do for k:=2 downto 1 do WriteLn(i*100+j*10+k); end."#
        ),
        &["222", "221", "212", "211", "122", "121", "112", "111"]
    );
}

#[test]
fn matrix_row_major_indices() {
    assert_eq!(
        run_pascal(
            r#"program T; var r,c:Integer; begin for r:=0 to 2 do for c:=0 to 2 do WriteLn(r*3+c); end."#
        ),
        &["0", "1", "2", "3", "4", "5", "6", "7", "8"]
    );
}

#[test]
fn diagonal_only_nested() {
    assert_eq!(
        run_pascal(
            r#"program T; var r,c:Integer; begin for r:=1 to 3 do for c:=1 to 3 do if r=c then WriteLn(r); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn triple_while_countdown_print() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c:Integer; begin a:=2; while a>0 do begin b:=2; while b>0 do begin c:=2; while c>0 do begin WriteLn(a*100+b*10+c); Dec(c); end; Dec(b); end; Dec(a); end; end."#
        ),
        &["222", "221", "212", "211", "122", "121", "112", "111"]
    );
}

#[test]
fn manual_step_by_three() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; while i<=9 do begin WriteLn(i); i:=i+3; end; end."#
        ),
        &["0", "3", "6", "9"]
    );
}

#[test]
fn break_from_inner_repeat() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin i:=0; repeat Inc(i); j:=0; repeat Inc(j); if j=2 then Break; until false; until i=3; WriteLn(i); WriteLn(j); end."#
        ),
        &["3", "2"]
    );
}

#[test]
fn search_break_on_product() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,found:Integer; begin found:=0; for i:=1 to 5 do if found=0 then for j:=1 to 5 do if (i*j=15) and (found=0) then begin found:=1; WriteLn(i); WriteLn(j); Break; end; end."#
        ),
        &["3", "5"]
    );
}

#[test]
fn pyramid_cell_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var r,c,n:Integer; begin n:=0; for r:=1 to 4 do for c:=1 to r do Inc(n); WriteLn(n); end."#
        ),
        &["10"]
    );
}

#[test]
fn break_outer_after_row_limit() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 4 do begin for j:=1 to 3 do WriteLn(i*10+j); if i=2 then Break; end; end."#
        ),
        &["11", "12", "13", "21", "22", "23"]
    );
}

#[test]
fn clock_nested_tick_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var h,m,s,c:Integer; begin c:=0; for h:=0 to 1 do for m:=0 to 1 do for s:=0 to 1 do Inc(c); WriteLn(c); end."#
        ),
        &["8"]
    );
}

#[test]
fn inner_for_shadow_print() {
    assert_eq!(
        run_pascal(
            r#"program T; var o,j:Integer; begin for o:=1 to 2 do begin for j:=5 to 6 do WriteLn(j); WriteLn('o'); end; end."#
        ),
        &["5", "6", "o", "5", "6", "o"]
    );
}

#[test]
fn repeat_flag_stop_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; stop:Boolean; begin stop:=false; i:=0; repeat Inc(i); j:=0; repeat Inc(j); if (i=2) and (j=2) then stop:=true; until stop or (j>=3); until stop; WriteLn(i); WriteLn(j); end."#
        ),
        &["2", "2"]
    );
}

#[test]
fn while_flag_leave_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; leave:Boolean; begin leave:=false; i:=0; while (i<5) and not leave do begin Inc(i); j:=0; while j<5 do begin Inc(j); if (i=3) and (j=2) then leave:=true; end; end; WriteLn(i); end."#
        ),
        &["3"]
    );
}

#[test]
fn mult_table_2x2() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=2 to 3 do for j:=2 to 3 do WriteLn(i*j); end."#
        ),
        &["4", "6", "6", "9"]
    );
}

#[test]
fn deep_sum_mod_three_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c,s:Integer; begin s:=0; for a:=1 to 3 do for b:=1 to 3 do for c:=1 to 3 do if (a+b+c) mod 3=0 then Inc(s); WriteLn(s); end."#
        ),
        &["9"]
    );
}

#[test]
fn labeled_break_sim_sum_six() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,k,stop:Integer; begin stop:=0; for i:=1 to 4 do if stop=0 then for j:=1 to 4 do if stop=0 then for k:=1 to 4 do if i+j+k=6 then begin WriteLn(i); WriteLn(j); WriteLn(k); stop:=1; end; end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn labeled_break_sim_sum_seven() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,k,stop:Integer; begin stop:=0; for i:=1 to 4 do if stop=0 then for j:=1 to 4 do if stop=0 then for k:=1 to 4 do if i+j+k=7 then begin WriteLn(i); WriteLn(j); WriteLn(k); stop:=1; end; end."#
        ),
        &["1", "2", "4"]
    );
}

#[test]
fn nested_repeat_count_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,n:Integer; begin n:=0; i:=0; repeat Inc(i); j:=0; repeat Inc(j); Inc(n); until j=2; until i=2; WriteLn(n); end."#
        ),
        &["6"]
    );
}

#[test]
fn for_downto_inside_for_to() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 2 do for j:=2 downto 1 do WriteLn(i*10+j); end."#
        ),
        &["12", "11", "22", "21"]
    );
}

#[test]
fn while_nested_sum_until() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,s:Integer; begin s:=0; i:=0; while i<3 do begin Inc(i); j:=0; while j<i do begin Inc(j); s:=s+1; end; end; WriteLn(s); end."#
        ),
        &["6"]
    );
}

#[test]
fn repeat_until_outer_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y:Integer; begin x:=0; repeat Inc(x); y:=0; repeat Inc(y); until y=2; until x=2; WriteLn(x); WriteLn(y); end."#
        ),
        &["2", "2"]
    );
}

#[test]
fn nested_break_first_match() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 4 do for j:=1 to 4 do if i*j=12 then begin WriteLn(i); WriteLn(j); Break; end; end."#
        ),
        &["3", "4"]
    );
}

#[test]
fn grid_skip_center() {
    assert_eq!(
        run_pascal(
            r#"program T; var r,c:Integer; begin for r:=1 to 3 do for c:=1 to 3 do if not ((r=2) and (c=2)) then WriteLn(r*10+c); end."#
        ),
        &["11", "12", "13", "21", "23", "31", "32", "33"]
    );
}

#[test]
fn outer_flag_after_inner_scan() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,hit:Integer; begin hit:=0; for i:=1 to 5 do begin if hit=0 then for j:=1 to 5 do if i+j=7 then hit:=i; end; WriteLn(hit); end."#
        ),
        &["2"]
    );
}

#[test]
fn triple_for_break_on_sum_ten() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c,stop:Integer; begin stop:=0; for a:=1 to 4 do if stop=0 then for b:=1 to 4 do if stop=0 then for c:=1 to 4 do if a+b+c=10 then begin WriteLn(a); WriteLn(b); WriteLn(c); stop:=1; end; end."#
        ),
        &["1", "3", "6"]
    );
}

#[test]
fn nested_while_find_pair() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; found:Boolean; begin found:=false; a:=1; while (a<=3) and not found do begin b:=1; while b<=3 do begin if a*a+b*b=10 then begin WriteLn(a); WriteLn(b); found:=true; Break; end; Inc(b); end; Inc(a); end; end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn for_repeat_write_rows() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; begin for i:=1 to 2 do begin j:=0; repeat Inc(j); WriteLn(i*10+j); until j=2; end; end."#
        ),
        &["11", "12", "21", "22"]
    );
}

#[test]
fn deep_nest_boolean_filter() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y,z,n:Integer; begin n:=0; for x:=0 to 1 do for y:=0 to 1 do for z:=0 to 1 do if (x+y+z)>=2 then Inc(n); WriteLn(n); end."#
        ),
        &["4"]
    );
}

#[test]
fn simulate_labeled_continue_skip() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,s:Integer; begin s:=0; for i:=1 to 3 do for j:=1 to 3 do begin if j=2 then begin end else s:=s+1; end; WriteLn(s); end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_loop_max_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j,m:Integer; begin m:=0; for i:=1 to 3 do for j:=1 to 3 do if i*j>m then m:=i*j; WriteLn(m); end."#
        ),
        &["9"]
    );
}

#[test]
fn break_outer_via_goto_flag_combo() {
    assert_eq!(
        run_pascal(
            r#"program T; var i,j:Integer; abort:Boolean; begin abort:=false; for i:=1 to 10 do begin if abort then Break; for j:=1 to 10 do if i*j=20 then begin WriteLn(i); WriteLn(j); abort:=true; Break; end; end; end."#
        ),
        &["4", "5"]
    );
}

#[test]
fn five_level_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c,d,e,n:Integer; begin n:=0; for a:=0 to 1 do for b:=0 to 1 do for c:=0 to 1 do for d:=0 to 1 do for e:=0 to 1 do Inc(n); WriteLn(n); end."#
        ),
        &["32"]
    );
}

