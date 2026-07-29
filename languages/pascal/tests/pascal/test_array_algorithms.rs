/// Sort and search algorithms on static arrays.
use super::helpers::run_pascal;

#[test]
fn linear_search_found_middle() {
    assert_eq!(
        run_pascal(
            r#"program T; function Find(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=-1; for i:=0 to High(a) do if a[i]=v then Result:=i; end; var arr:array[0..4] of Integer; begin arr[0]:=3; arr[1]:=7; arr[2]:=9; arr[3]:=2; arr[4]:=5; WriteLn(Find(arr,9)); end."#
        ),
        &["2"]
    );
}

#[test]
fn linear_search_not_found() {
    assert_eq!(
        run_pascal(
            r#"program T; function Find(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=-1; for i:=0 to High(a) do if a[i]=v then Result:=i; end; var arr:array[0..2] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=3; WriteLn(Find(arr,4)); end."#
        ),
        &["-1"]
    );
}

#[test]
fn bubble_sort_ascending_small() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Sort(var a:array of Integer); var i,j,t,n:Integer; begin n:=High(a); for i:=0 to n do for j:=0 to n-1 do if a[j]>a[j+1] then begin t:=a[j]; a[j]:=a[j+1]; a[j+1]:=t; end; end; var arr:array[0..3] of Integer; begin arr[0]:=4; arr[1]:=1; arr[2]:=3; arr[3]:=2; Sort(arr); WriteLn(arr[0]); WriteLn(arr[3]); end."#
        ),
        &["1", "4"]
    );
}

#[test]
fn selection_sort_picks_min() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure SelSort(var a:array of Integer); var i,j,m,t:Integer; begin for i:=0 to High(a) do begin m:=i; for j:=i+1 to High(a) do if a[j]<a[m] then m:=j; t:=a[i]; a[i]:=a[m]; a[m]:=t; end; end; var arr:array[0..2] of Integer; begin arr[0]:=5; arr[1]:=1; arr[2]:=3; SelSort(arr); WriteLn(arr[0]); end."#
        ),
        &["1"]
    );
}

#[test]
fn insertion_sort_shift_right() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure InsSort(var a:array of Integer); var i,j,k:Integer; begin for i:=1 to High(a) do begin k:=a[i]; j:=i-1; while (j>=0) and (a[j]>k) do begin a[j+1]:=a[j]; Dec(j); end; a[j+1]:=k; end; end; var arr:array[0..3] of Integer; begin arr[0]:=3; arr[1]:=1; arr[2]:=4; arr[3]:=2; InsSort(arr); WriteLn(arr[1]); WriteLn(arr[2]); end."#
        ),
        &["2", "3"]
    );
}

#[test]
fn array_sum_static() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sum(const a:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=0 to High(a) do Result:=Result+a[i]; end; var arr:array[0..3] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=3; arr[3]:=4; WriteLn(Sum(arr)); end."#
        ),
        &["10"]
    );
}

#[test]
fn array_max_element() {
    assert_eq!(
        run_pascal(
            r#"program T; function MaxVal(const a:array of Integer):Integer; var i:Integer; begin Result:=a[0]; for i:=1 to High(a) do if a[i]>Result then Result:=a[i]; end; var arr:array[0..2] of Integer; begin arr[0]:=2; arr[1]:=9; arr[2]:=5; WriteLn(MaxVal(arr)); end."#
        ),
        &["9"]
    );
}

#[test]
fn array_min_element() {
    assert_eq!(
        run_pascal(
            r#"program T; function MinVal(const a:array of Integer):Integer; var i:Integer; begin Result:=a[0]; for i:=1 to High(a) do if a[i]<Result then Result:=a[i]; end; var arr:array[0..2] of Integer; begin arr[0]:=2; arr[1]:=9; arr[2]:=5; WriteLn(MinVal(arr)); end."#
        ),
        &["2"]
    );
}

#[test]
fn binary_search_sorted_hit() {
    assert_eq!(
        run_pascal(
            r#"program T; function BSearch(const a:array of Integer; v:Integer):Integer; var lo,hi,m:Integer; begin lo:=0; hi:=High(a); Result:=-1; while lo<=hi do begin m:=(lo+hi) div 2; if a[m]=v then begin Result:=m; Break; end else if a[m]<v then lo:=m+1 else hi:=m-1; end; end; var arr:array[0..4] of Integer; begin arr[0]:=1; arr[1]:=3; arr[2]:=5; arr[3]:=7; arr[4]:=9; WriteLn(BSearch(arr,7)); end."#
        ),
        &["3"]
    );
}

#[test]
fn binary_search_sorted_miss() {
    assert_eq!(
        run_pascal(
            r#"program T; function BSearch(const a:array of Integer; v:Integer):Integer; var lo,hi,m:Integer; begin lo:=0; hi:=High(a); Result:=-1; while lo<=hi do begin m:=(lo+hi) div 2; if a[m]=v then begin Result:=m; Break; end else if a[m]<v then lo:=m+1 else hi:=m-1; end; end; var arr:array[0..3] of Integer; begin arr[0]:=2; arr[1]:=4; arr[2]:=6; arr[3]:=8; WriteLn(BSearch(arr,5)); end."#
        ),
        &["-1"]
    );
}

#[test]
fn reverse_array_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Rev(var a:array of Integer); var i,j,t:Integer; begin j:=High(a); for i:=0 to j div 2 do begin t:=a[i]; a[i]:=a[j-i]; a[j-i]:=t; end; end; var arr:array[0..3] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=3; arr[3]:=4; Rev(arr); WriteLn(arr[0]); WriteLn(arr[3]); end."#
        ),
        &["4", "1"]
    );
}

#[test]
fn count_occurrences_value() {
    assert_eq!(
        run_pascal(
            r#"program T; function Count(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=0; for i:=0 to High(a) do if a[i]=v then Inc(Result); end; var arr:array[0..4] of Integer; begin arr[0]:=2; arr[1]:=2; arr[2]:=3; arr[3]:=2; arr[4]:=1; WriteLn(Count(arr,2)); end."#
        ),
        &["3"]
    );
}

#[test]
fn fill_array_with_constant() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Fill(var a:array of Integer; v:Integer); var i:Integer; begin for i:=0 to High(a) do a[i]:=v; end; var arr:array[0..2] of Integer; begin Fill(arr,7); WriteLn(arr[1]); end."#
        ),
        &["7"]
    );
}

#[test]
fn copy_array_segment_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var src:array[1..5] of Integer; i,s:Integer; begin for i:=1 to 5 do src[i]:=i; s:=0; for i:=2 to 4 do s:=s+src[i]; WriteLn(s); end."#
        ),
        &["9"]
    );
}

#[test]
fn static_char_array_join() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[1..3] of Char; s:string; begin a[1]:='a'; a[2]:='b'; a[3]:='c'; s:=a[1]+a[2]+a[3]; WriteLn(s); end."#
        ),
        &["abc"]
    );
}

#[test]
fn one_based_array_bounds_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[1..4] of Integer; i,s:Integer; begin for i:=1 to 4 do a[i]:=i; s:=0; for i:=Low(a) to High(a) do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn zero_based_array_index_first() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..2] of Integer; begin a[0]:=11; WriteLn(a[0]); end."#
        ),
        &["11"]
    );
}

#[test]
fn negative_lower_bound_array() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[-1..1] of Integer; begin a[-1]:=2; a[1]:=3; WriteLn(a[-1]+a[1]); end."#
        ),
        &["5"]
    );
}

#[test]
fn bubble_sort_already_sorted() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Sort(var a:array of Integer); var i,j,t,n:Integer; begin n:=High(a); for i:=0 to n do for j:=0 to n-1 do if a[j]>a[j+1] then begin t:=a[j]; a[j]:=a[j+1]; a[j+1]:=t; end; end; var arr:array[0..2] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=3; Sort(arr); WriteLn(arr[2]); end."#
        ),
        &["3"]
    );
}

#[test]
fn find_first_greater_than() {
    assert_eq!(
        run_pascal(
            r#"program T; function FirstGT(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=-1; for i:=0 to High(a) do if a[i]>v then begin Result:=i; Break; end; end; var arr:array[0..3] of Integer; begin arr[0]:=1; arr[1]:=3; arr[2]:=3; arr[3]:=6; WriteLn(FirstGT(arr,3)); end."#
        ),
        &["3"]
    );
}

#[test]
fn array_average_integer() {
    assert_eq!(
        run_pascal(
            r#"program T; function Avg(const a:array of Integer):Integer; var i,s:Integer; begin s:=0; for i:=0 to High(a) do s:=s+a[i]; Result:=s div (High(a)+1); end; var arr:array[0..3] of Integer; begin arr[0]:=2; arr[1]:=4; arr[2]:=6; arr[3]:=8; WriteLn(Avg(arr)); end."#
        ),
        &["5"]
    );
}

#[test]
fn shift_left_array_manual() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..3] of Integer; i:Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; for i:=0 to 2 do a[i]:=a[i+1]; a[3]:=0; WriteLn(a[0]); WriteLn(a[3]); end."#
        ),
        &["2", "0"]
    );
}

#[test]
fn duplicate_check_has_pair() {
    assert_eq!(
        run_pascal(
            r#"program T; function HasDup(const a:array of Integer):Boolean; var i,j:Integer; begin Result:=false; for i:=0 to High(a) do for j:=i+1 to High(a) do if a[i]=a[j] then Result:=true; end; var arr:array[0..3] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=2; arr[3]:=3; WriteLn(HasDup(arr)); end."#
        ),
        &["True"]
    );
}

#[test]
fn duplicate_check_all_unique() {
    assert_eq!(
        run_pascal(
            r#"program T; function HasDup(const a:array of Integer):Boolean; var i,j:Integer; begin Result:=false; for i:=0 to High(a) do for j:=i+1 to High(a) do if a[i]=a[j] then Result:=true; end; var arr:array[0..2] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=3; WriteLn(HasDup(arr)); end."#
        ),
        &["False"]
    );
}

#[test]
fn counting_sort_small_range() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,c:array[0..4] of Integer; i,v:Integer; begin a[0]:=3; a[1]:=1; a[2]:=2; a[3]:=1; a[4]:=3; for i:=0 to 4 do c[i]:=0; for i:=0 to 4 do Inc(c[a[i]]); v:=0; for i:=0 to 4 do while c[i]>0 do begin a[v]:=i; Dec(c[i]); Inc(v); end; WriteLn(a[0]); WriteLn(a[4]); end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn binary_search_first_element() {
    assert_eq!(
        run_pascal(
            r#"program T; function BSearch(const a:array of Integer; v:Integer):Integer; var lo,hi,m:Integer; begin lo:=0; hi:=High(a); Result:=-1; while lo<=hi do begin m:=(lo+hi) div 2; if a[m]=v then begin Result:=m; Break; end else if a[m]<v then lo:=m+1 else hi:=m-1; end; end; var arr:array[0..2] of Integer; begin arr[0]:=2; arr[1]:=4; arr[2]:=6; WriteLn(BSearch(arr,2)); end."#
        ),
        &["0"]
    );
}

#[test]
fn binary_search_last_element() {
    assert_eq!(
        run_pascal(
            r#"program T; function BSearch(const a:array of Integer; v:Integer):Integer; var lo,hi,m:Integer; begin lo:=0; hi:=High(a); Result:=-1; while lo<=hi do begin m:=(lo+hi) div 2; if a[m]=v then begin Result:=m; Break; end else if a[m]<v then lo:=m+1 else hi:=m-1; end; end; var arr:array[0..2] of Integer; begin arr[0]:=2; arr[1]:=4; arr[2]:=6; WriteLn(BSearch(arr,6)); end."#
        ),
        &["2"]
    );
}

#[test]
fn array_index_of_min() {
    assert_eq!(
        run_pascal(
            r#"program T; function IdxMin(const a:array of Integer):Integer; var i,m:Integer; begin m:=0; for i:=1 to High(a) do if a[i]<a[m] then m:=i; Result:=m; end; var arr:array[0..3] of Integer; begin arr[0]:=5; arr[1]:=2; arr[2]:=8; arr[3]:=1; WriteLn(IdxMin(arr)); end."#
        ),
        &["3"]
    );
}

#[test]
fn array_index_of_max() {
    assert_eq!(
        run_pascal(
            r#"program T; function IdxMax(const a:array of Integer):Integer; var i,m:Integer; begin m:=0; for i:=1 to High(a) do if a[i]>a[m] then m:=i; Result:=m; end; var arr:array[0..3] of Integer; begin arr[0]:=5; arr[1]:=2; arr[2]:=8; arr[3]:=1; WriteLn(IdxMax(arr)); end."#
        ),
        &["2"]
    );
}

#[test]
fn partial_sort_first_two_bubble() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..3] of Integer; i,j,t:Integer; begin a[0]:=4; a[1]:=2; a[2]:=3; a[3]:=1; for i:=0 to 1 do for j:=0 to 3-i-1 do if a[j]>a[j+1] then begin t:=a[j]; a[j]:=a[j+1]; a[j+1]:=t; end; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn static_bool_array_count_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..3] of Boolean; i,c:Integer; begin a[0]:=true; a[1]:=false; a[2]:=true; a[3]:=true; c:=0; for i:=0 to 3 do if a[i] then Inc(c); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn array_rotate_right_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..3] of Integer; t:Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; t:=a[3]; a[3]:=a[2]; a[2]:=a[1]; a[1]:=a[0]; a[0]:=t; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["4", "1"]
    );
}

#[test]
fn two_array_dot_product() {
    assert_eq!(
        run_pascal(
            r#"program T; function Dot(const u,v:array of Integer):Integer; var i:Integer; begin Result:=0; for i:=0 to High(u) do Result:=Result+u[i]*v[i]; end; var a,b:array[0..2] of Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; b[0]:=4; b[1]:=5; b[2]:=6; WriteLn(Dot(a,b)); end."#
        ),
        &["32"]
    );
}

#[test]
fn array_all_equal_check() {
    assert_eq!(
        run_pascal(
            r#"program T; function AllEq(const a:array of Integer; v:Integer):Boolean; var i:Integer; begin Result:=true; for i:=0 to High(a) do if a[i]<>v then Result:=false; end; var arr:array[0..2] of Integer; begin arr[0]:=5; arr[1]:=5; arr[2]:=5; WriteLn(AllEq(arr,5)); end."#
        ),
        &["True"]
    );
}

#[test]
fn array_is_sorted_ascending() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sorted(const a:array of Integer):Boolean; var i:Integer; begin Result:=true; for i:=0 to High(a)-1 do if a[i]>a[i+1] then Result:=false; end; var arr:array[0..3] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=2; arr[3]:=4; WriteLn(Sorted(arr)); end."#
        ),
        &["True"]
    );
}

#[test]
fn array_is_not_sorted() {
    assert_eq!(
        run_pascal(
            r#"program T; function Sorted(const a:array of Integer):Boolean; var i:Integer; begin Result:=true; for i:=0 to High(a)-1 do if a[i]>a[i+1] then Result:=false; end; var arr:array[0..2] of Integer; begin arr[0]:=1; arr[1]:=3; arr[2]:=2; WriteLn(Sorted(arr)); end."#
        ),
        &["False"]
    );
}

#[test]
fn find_last_occurrence_linear() {
    assert_eq!(
        run_pascal(
            r#"program T; function Last(const a:array of Integer; v:Integer):Integer; var i:Integer; begin Result:=-1; for i:=0 to High(a) do if a[i]=v then Result:=i; end; var arr:array[0..4] of Integer; begin arr[0]:=1; arr[1]:=2; arr[2]:=2; arr[3]:=2; arr[4]:=3; WriteLn(Last(arr,2)); end."#
        ),
        &["3"]
    );
}

#[test]
fn array_median_of_three_values() {
    assert_eq!(
        run_pascal(
            r#"program T; function Med3(a,b,c:Integer):Integer; begin if ((a<=b) and (b<=c)) or ((c<=b) and (b<=a)) then Result:=b else if ((b<=a) and (a<=c)) or ((c<=a) and (a<=b)) then Result:=a else Result:=c; end; begin WriteLn(Med3(3,1,2)); end."#
        ),
        &["2"]
    );
}

#[test]
fn static_array_init_literal_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..2] of Integer = (2,4,6); s,i:Integer; begin s:=0; for i:=0 to 2 do s:=s+a[i]; WriteLn(s); end."#
        ),
        &["12"]
    );
}

#[test]
fn partition_pivot_count_less() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..4] of Integer; i,p,c:Integer; begin a[0]:=5; a[1]:=2; a[2]:=8; a[3]:=1; a[4]:=6; p:=5; c:=0; for i:=0 to 4 do if a[i]<p then Inc(c); WriteLn(c); end."#
        ),
        &["2"]
    );
}

#[test]
fn array_swap_two_indices() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..2] of Integer; t:Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; t:=a[0]; a[0]:=a[2]; a[2]:=t; WriteLn(a[0]); WriteLn(a[2]); end."#
        ),
        &["3", "1"]
    );
}

#[test]
fn gnome_sort_tiny() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure GSort(var a:array of Integer); var i:Integer; begin i:=1; while i<=High(a) do begin if (i=0) or (a[i-1]<=a[i]) then Inc(i) else begin a[i-1]:=a[i-1] xor a[i]; a[i]:=a[i-1] xor a[i]; a[i-1]:=a[i-1] xor a[i]; Dec(i); end; end; end; var arr:array[0..3] of Integer; begin arr[0]:=3; arr[1]:=1; arr[2]:=4; arr[3]:=2; GSort(arr); WriteLn(arr[0]); WriteLn(arr[3]); end."#
        ),
        &["1", "4"]
    );
}
