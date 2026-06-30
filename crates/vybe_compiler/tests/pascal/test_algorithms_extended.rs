/// Sorting, searching, GCD/LCM, and classic algorithms — distinct from test_algorithms4.rs.
use super::helpers::run_pascal;

#[test]
fn euclidean_gcd_iterative() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
var tmp: Integer;
begin
  while b <> 0 do begin tmp := a mod b; a := b; b := tmp; end;
  Result := a;
end;
begin WriteLn(GCD(48, 18)); end."#
        ),
        &["6"]
    );
}

#[test]
fn euclidean_gcd_recursive() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
begin WriteLn(GCD(100, 35)); end."#
        ),
        &["5"]
    );
}

#[test]
fn lcm_from_gcd_formula() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
function LCM(a, b: Integer): Integer;
begin Result := (a div GCD(a, b)) * b; end;
begin WriteLn(LCM(12, 18)); end."#
        ),
        &["36"]
    );
}

#[test]
fn lcm_three_numbers() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
function LCM(a, b: Integer): Integer;
begin Result := (a div GCD(a, b)) * b; end;
begin WriteLn(LCM(LCM(4, 6), 10)); end."#
        ),
        &["60"]
    );
}

#[test]
fn cocktail_sort_shaker() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    i, lo, hi, tmp: Integer;
    swapped: Boolean;
begin
  a[0]:=5; a[1]:=1; a[2]:=4; a[3]:=2; a[4]:=3;
  lo := 0; hi := 4;
  while lo < hi do
  begin
    swapped := False;
    for i := lo to hi - 1 do
      if a[i] > a[i+1] then begin tmp:=a[i]; a[i]:=a[i+1]; a[i+1]:=tmp; swapped:=True; end;
    if not swapped then Break;
    swapped := False;
    Dec(hi);
    for i := hi downto lo + 1 do
      if a[i] < a[i-1] then begin tmp:=a[i]; a[i]:=a[i-1]; a[i-1]:=tmp; swapped:=True; end;
    if not swapped then Break;
    Inc(lo);
  end;
  for i := 0 to 4 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["1 2 3 4 5 "]
    );
}

#[test]
fn comb_sort_gap_shrink() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    i, gap, n, tmp: Integer;
    swapped: Boolean;
begin
  a[0]:=8; a[1]:=4; a[2]:=1; a[3]:=7; a[4]:=3; a[5]:=2;
  n := 6; gap := n;
  repeat
    gap := (gap * 10) div 13;
    if gap < 1 then gap := 1;
    swapped := False;
    for i := 0 to n - gap - 1 do
      if a[i] > a[i+gap] then begin tmp:=a[i]; a[i]:=a[i+gap]; a[i+gap]:=tmp; swapped:=True; end;
  until (gap = 1) and not swapped;
  for i := 0 to 5 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["1 2 3 4 7 8 "]
    );
}

#[test]
fn selection_sort_minimum_swap() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    i, j, minIdx, tmp: Integer;
begin
  a[0]:=64; a[1]:=25; a[2]:=12; a[3]:=22; a[4]:=11;
  for i := 0 to 3 do
  begin
    minIdx := i;
    for j := i + 1 to 4 do if a[j] < a[minIdx] then minIdx := j;
    tmp := a[i]; a[i] := a[minIdx]; a[minIdx] := tmp;
  end;
  for i := 0 to 4 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["11 12 22 25 64 "]
    );
}

#[test]
fn bubble_sort_early_exit() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..3] of Integer;
    i, j, tmp: Integer;
    swapped: Boolean;
begin
  a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4;
  for i := 0 to 2 do
  begin
    swapped := False;
    for j := 0 to 2 - i do
      if a[j] > a[j+1] then begin tmp:=a[j]; a[j]:=a[j+1]; a[j+1]:=tmp; swapped:=True; end;
    if not swapped then Break;
  end;
  WriteLn(a[0]);
end."#
        ),
        &["1"]
    );
}

#[test]
fn quicksort_partition_lomuto() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
procedure QSort(lo, hi: Integer);
var i, j, pivot, tmp: Integer;
begin
  if lo >= hi then Exit;
  pivot := a[hi]; i := lo;
  for j := lo to hi - 1 do
    if a[j] <= pivot then begin tmp:=a[i]; a[i]:=a[j]; a[j]:=tmp; Inc(i); end;
  tmp := a[i]; a[i] := a[hi]; a[hi] := tmp;
  QSort(lo, i - 1); QSort(i + 1, hi);
end;
var i: Integer;
begin
  a[0]:=10; a[1]:=7; a[2]:=8; a[3]:=9; a[4]:=1; a[5]:=5;
  QSort(0, 5);
  for i := 0 to 5 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["1 5 7 8 9 10 "]
    );
}

#[test]
fn mergesort_two_halves() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    tmp: array[0..5] of Integer;
procedure Merge(lo, mid, hi: Integer);
var i, j, k: Integer;
begin
  i := lo; j := mid + 1; k := lo;
  while (i <= mid) and (j <= hi) do
    if a[i] <= a[j] then begin tmp[k]:=a[i]; Inc(i); end else begin tmp[k]:=a[j]; Inc(j); end;
    Inc(k);
  while i <= mid do begin tmp[k]:=a[i]; Inc(i); Inc(k); end;
  while j <= hi do begin tmp[k]:=a[j]; Inc(j); Inc(k); end;
  for i := lo to hi do a[i] := tmp[i];
end;
procedure MSort(lo, hi: Integer);
var mid: Integer;
begin
  if lo >= hi then Exit;
  mid := (lo + hi) div 2;
  MSort(lo, mid); MSort(mid + 1, hi); Merge(lo, mid, hi);
end;
var i: Integer;
begin
  a[0]:=38; a[1]:=27; a[2]:=43; a[3]:=3; a[4]:=9; a[5]:=82;
  MSort(0, 5);
  for i := 0 to 5 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["3 9 27 38 43 82 "]
    );
}

#[test]
fn heap_sort_build_and_extract() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
procedure Heapify(n, i: Integer);
var largest, l, r, tmp: Integer;
begin
  largest := i; l := 2*i+1; r := 2*i+2;
  if (l < n) and (a[l] > a[largest]) then largest := l;
  if (r < n) and (a[r] > a[largest]) then largest := r;
  if largest <> i then begin tmp:=a[i]; a[i]:=a[largest]; a[largest]:=tmp; Heapify(n, largest); end;
end;
var i, n, tmp: Integer;
begin
  a[0]:=4; a[1]:=10; a[2]:=3; a[3]:=5; a[4]:=1;
  n := 5;
  for i := n div 2 downto 0 do Heapify(n, i);
  for i := n - 1 downto 1 do
  begin tmp:=a[0]; a[0]:=a[i]; a[i]:=tmp; Heapify(i, 0); end;
  for i := 0 to 4 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["1 3 4 5 10 "]
    );
}

#[test]
fn binary_search_recursive() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..6] of Integer;
function BSearch(lo, hi, key: Integer): Integer;
var mid: Integer;
begin
  if lo > hi then Result := -1
  else begin
    mid := (lo + hi) div 2;
    if a[mid] = key then Result := mid
    else if a[mid] < key then Result := BSearch(mid + 1, hi, key)
    else Result := BSearch(lo, mid - 1, key);
  end;
end;
begin
  a[0]:=2; a[1]:=5; a[2]:=8; a[3]:=12; a[4]:=16; a[5]:=23; a[6]:=38;
  WriteLn(BSearch(0, 6, 16));
end."#
        ),
        &["4"]
    );
}

#[test]
fn binary_search_first_occurrence() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
function FindFirst(key: Integer): Integer;
var lo, hi, mid: Integer;
begin
  lo := 0; hi := 5; Result := -1;
  while lo <= hi do
  begin
    mid := (lo + hi) div 2;
    if a[mid] = key then begin Result := mid; hi := mid - 1; end
    else if a[mid] < key then lo := mid + 1 else hi := mid - 1;
  end;
end;
begin
  a[0]:=1; a[1]:=2; a[2]:=2; a[3]:=2; a[4]:=3; a[5]:=4;
  WriteLn(FindFirst(2));
end."#
        ),
        &["1"]
    );
}

#[test]
fn linear_search_sentinel() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    i, key: Integer;
begin
  a[0]:=3; a[1]:=7; a[2]:=1; a[3]:=9; a[4]:=4;
  key := 9;
  i := 0;
  while a[i] <> key do Inc(i);
  WriteLn(i);
end."#
        ),
        &["3"]
    );
}

#[test]
fn jump_search_on_sorted_array() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..9] of Integer;
    n, step, prev, i, key: Integer;
begin
  for i := 0 to 9 do a[i] := i * 2;
  n := 10; key := 14; step := 3; prev := 0;
  while a[prev] < key do
  begin prev := step; step := step + 3; if prev >= n then Break; end;
  i := prev - 3;
  if i < 0 then i := 0;
  while (i < n) and (a[i] < key) do Inc(i);
  if (i < n) and (a[i] = key) then WriteLn(i) else WriteLn(-1);
end."#
        ),
        &["7"]
    );
}

#[test]
fn interpolation_search_estimate() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    lo, hi, pos, key: Integer;
begin
  a[0]:=10; a[1]:=20; a[2]:=30; a[3]:=40; a[4]:=50;
  key := 30; lo := 0; hi := 4; pos := -1;
  while (lo <= hi) and (key >= a[lo]) and (key <= a[hi]) do
  begin
    if lo = hi then begin if a[lo] = key then pos := lo; Break; end;
    pos := lo + ((key - a[lo]) * (hi - lo)) div (a[hi] - a[lo]);
    if a[pos] = key then Break else if a[pos] < key then lo := pos + 1 else hi := pos - 1;
  end;
  WriteLn(pos);
end."#
        ),
        &["2"]
    );
}

#[test]
fn ternary_search_maximum_unimodal() {
    assert_eq!(
        run_pascal(
            r#"program T;
function F(x: Integer): Integer;
begin Result := -(x - 5) * (x - 5) + 25; end;
var lo, hi, m1, m2: Integer;
begin
  lo := 0; hi := 10;
  while hi - lo > 2 do
  begin
    m1 := lo + (hi - lo) div 3;
    m2 := hi - (hi - lo) div 3;
    if F(m1) < F(m2) then lo := m1 else hi := m2;
  end;
  WriteLn(lo);
end."#
        ),
        &["5"]
    );
}

#[test]
fn fibonacci_iterative() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a, b, i, n, tmp: Integer;
begin
  n := 10; a := 0; b := 1;
  for i := 1 to n do begin tmp := a + b; a := b; b := tmp; end;
  WriteLn(b);
end."#
        ),
        &["89"]
    );
}

#[test]
fn factorial_iterative() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i, n, f: Integer;
begin
  n := 6; f := 1;
  for i := 2 to n do f := f * i;
  WriteLn(f);
end."#
        ),
        &["720"]
    );
}

#[test]
fn power_exponentiation_by_squaring() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Pow(base, exp: Integer): Integer;
var result: Integer;
begin
  result := 1;
  while exp > 0 do
  begin
    if (exp mod 2) = 1 then result := result * base;
    base := base * base;
    exp := exp div 2;
  end;
  Result := result;
end;
begin WriteLn(Pow(2, 10)); end."#
        ),
        &["1024"]
    );
}

#[test]
fn sieve_eratosthenes_count_primes_under_30() {
    assert_eq!(
        run_pascal(
            r#"program T;
var isPrime: array[0..29] of Boolean;
    i, j, count: Integer;
begin
  for i := 0 to 29 do isPrime[i] := True;
  isPrime[0] := False; isPrime[1] := False;
  for i := 2 to 29 do
    if isPrime[i] then
      for j := i * 2 to 29 do isPrime[j] := False;
  count := 0;
  for i := 2 to 29 do if isPrime[i] then Inc(count);
  WriteLn(count);
end."#
        ),
        &["10"]
    );
}

#[test]
fn reverse_array_in_place() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    i, lo, hi, tmp: Integer;
begin
  a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; a[4]:=5;
  lo := 0; hi := 4;
  while lo < hi do begin tmp:=a[lo]; a[lo]:=a[hi]; a[hi]:=tmp; Inc(lo); Dec(hi); end;
  for i := 0 to 4 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["5 4 3 2 1 "]
    );
}

#[test]
fn rotate_array_right_by_k() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    k, i, n, tmp: Integer;
begin
  a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; a[4]:=5;
  n := 5; k := 2 mod n;
  for i := 1 to k do begin tmp:=a[n-1]; a[3]:=a[2]; a[2]:=a[1]; a[1]:=a[0]; a[0]:=tmp; end;
  for i := 0 to 4 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["4 5 1 2 3 "]
    );
}

#[test]
fn dutch_flag_three_way_partition() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    lo, mid, hi, tmp: Integer;
begin
  a[0]:=2; a[1]:=0; a[2]:=2; a[3]:=1; a[4]:=0; a[5]:=1;
  lo := 0; mid := 0; hi := 5;
  while mid <= hi do
  begin
    if a[mid] = 0 then begin tmp:=a[lo]; a[lo]:=a[mid]; a[mid]:=tmp; Inc(lo); Inc(mid); end
    else if a[mid] = 1 then Inc(mid)
    else begin tmp:=a[mid]; a[mid]:=a[hi]; a[hi]:=tmp; Dec(hi); end;
  end;
  WriteLn(a[0]); WriteLn(a[5]);
end."#
        ),
        &["0", "2"]
    );
}

#[test]
fn kadane_max_subarray_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    i, cur, best: Integer;
begin
  a[0]:=-2; a[1]:=1; a[2]:=-3; a[3]:=4; a[4]:=-1; a[5]:=2;
  cur := 0; best := a[0];
  for i := 0 to 5 do
  begin cur := cur + a[i]; if cur > best then best := cur; if cur < 0 then cur := 0; end;
  WriteLn(best);
end."#
        ),
        &["5"]
    );
}

#[test]
fn dutch_national_flag_count_zeros() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    i, zeros: Integer;
begin
  a[0]:=1; a[1]:=0; a[2]:=1; a[3]:=0; a[4]:=1;
  zeros := 0;
  for i := 0 to 4 do if a[i] = 0 then Inc(zeros);
  WriteLn(zeros);
end."#
        ),
        &["2"]
    );
}

#[test]
fn binary_search_lower_bound() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
function LowerBound(key: Integer): Integer;
var lo, hi, mid: Integer;
begin
  lo := 0; hi := 5;
  while lo < hi do
  begin mid := (lo + hi) div 2; if a[mid] < key then lo := mid + 1 else hi := mid; end;
  Result := lo;
end;
begin
  a[0]:=1; a[1]:=2; a[2]:=4; a[3]:=4; a[4]:=5; a[5]:=6;
  WriteLn(LowerBound(4));
end."#
        ),
        &["2"]
    );
}

#[test]
fn gcd_of_three_numbers() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
begin WriteLn(GCD(GCD(54, 24), 18)); end."#
        ),
        &["6"]
    );
}

#[test]
fn lcm_pair_coprime() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
function LCM(a, b: Integer): Integer;
begin Result := (a div GCD(a, b)) * b; end;
begin WriteLn(LCM(7, 9)); end."#
        ),
        &["63"]
    );
}

#[test]
fn cycle_sort_min_swaps() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..3] of Integer;
    i, j, tmp: Integer;
begin
  a[0]:=3; a[1]:=1; a[2]:=2; a[3]:=0;
  for i := 0 to 2 do
  begin
    j := i;
    while a[j] <> i do Inc(j);
    if j <> i then begin tmp:=a[i]; a[i]:=a[j]; a[j]:=tmp; end;
  end;
  for i := 0 to 3 do Write(IntToStr(a[i]) + ' ');
  WriteLn('');
end."#
        ),
        &["0 1 2 3 "]
    );
}

#[test]
fn counting_sort_negative_shifted() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
    counts: array[0..4] of Integer;
    i, j, idx, minV, offset: Integer;
begin
  a[0]:=-1; a[1]:=0; a[2]:=-2; a[3]:=1; a[4]:=0;
  minV := -2; offset := -minV;
  for i := 0 to 4 do Inc(counts[a[i] + offset]);
  idx := 0;
  for i := 0 to 4 do
    for j := 1 to counts[i] do begin a[idx] := i + minV; Inc(idx); end;
  WriteLn(a[2]);
end."#
        ),
        &["0"]
    );
}

#[test]
fn stack_based_dfs_path_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
var adj: array[0..2, 0..1] of Integer;
    visited: array[0..2] of Boolean;
    stack: array[0..2] of Integer;
    top, node, sum: Integer;
begin
  adj[0,0]:=1; adj[0,1]:=2;
  adj[1,0]:=2; adj[1,1]:=-1;
  adj[2,0]:=-1; adj[2,1]:=-1;
  for node := 0 to 2 do visited[node] := False;
  top := 0; stack[0] := 0; visited[0] := True; sum := 0;
  while top >= 0 do
  begin
    node := stack[top]; Dec(top);
    sum := sum + node;
    if adj[node,0] >= 0 then if not visited[adj[node,0]] then begin Inc(top); stack[top]:=adj[node,0]; visited[adj[node,0]]:=True; end;
    if adj[node,1] >= 0 then if not visited[adj[node,1]] then begin Inc(top); stack[top]:=adj[node,1]; visited[adj[node,1]]:=True; end;
  end;
  WriteLn(sum);
end."#
        ),
        &["3"]
    );
}

#[test]
fn queue_bfs_level_order() {
    assert_eq!(
        run_pascal(
            r#"program T;
var q: array[0..3] of Integer;
    head, tail, x: Integer;
begin
  head := 0; tail := 0;
  q[tail] := 1; Inc(tail);
  q[tail] := 2; Inc(tail);
  q[tail] := 3; Inc(tail);
  while head < tail do
  begin x := q[head]; Inc(head); WriteLn(x); end;
end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn moore_majority_vote() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    i, cand, count: Integer;
begin
  a[0]:=2; a[1]:=2; a[2]:=1; a[3]:=2; a[4]:=1; a[5]:=2;
  cand := a[0]; count := 1;
  for i := 1 to 5 do
    if count = 0 then begin cand := a[i]; count := 1; end
    else if a[i] = cand then Inc(count) else Dec(count);
  WriteLn(cand);
end."#
        ),
        &["2"]
    );
}

#[test]
fn binary_search_upper_bound() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..4] of Integer;
function UpperBound(key: Integer): Integer;
var lo, hi, mid: Integer;
begin
  lo := 0; hi := 4;
  while lo < hi do
  begin mid := (lo + hi) div 2; if a[mid] <= key then lo := mid + 1 else hi := mid; end;
  Result := lo;
end;
begin
  a[0]:=1; a[1]:=2; a[2]:=2; a[3]:=3; a[4]:=4;
  WriteLn(UpperBound(2));
end."#
        ),
        &["3"]
    );
}

#[test]
fn extended_gcd_bezout() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
var g: Integer;
begin g := GCD(240, 46); WriteLn(g); end."#
        ),
        &["2"]
    );
}

#[test]
fn digit_sum_of_number() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n, s: Integer;
begin
  n := 12345; s := 0;
  while n > 0 do begin s := s + n mod 10; n := n div 10; end;
  WriteLn(s);
end."#
        ),
        &["15"]
    );
}

#[test]
fn palindrome_check_integer() {
    assert_eq!(
        run_pascal(
            r#"program T;
var n, rev, orig: Integer;
begin
  orig := 121; n := orig; rev := 0;
  while n > 0 do begin rev := rev * 10 + n mod 10; n := n div 10; end;
  if rev = orig then WriteLn('yes') else WriteLn('no');
end."#
        ),
        &["yes"]
    );
}

#[test]
fn coin_change_min_coins() {
    assert_eq!(
        run_pascal(
            r#"program T;
var coins: array[0..2] of Integer;
    amount, i, count, remain: Integer;
begin
  coins[0]:=1; coins[1]:=5; coins[2]:=11;
  amount := 15; count := 0; remain := amount;
  for i := 2 downto 0 do
    while remain >= coins[i] do begin remain := remain - coins[i]; Inc(count); end;
  WriteLn(count);
end."#
        ),
        &["3"]
    );
}

#[test]
fn sliding_window_max_sum() {
    assert_eq!(
        run_pascal(
            r#"program T;
var a: array[0..5] of Integer;
    i, k, winSum, best: Integer;
begin
  a[0]:=1; a[1]:=4; a[2]:=2; a[3]:=8; a[4]:=5; a[5]:=7;
  k := 3; winSum := 0;
  for i := 0 to k - 1 do winSum := winSum + a[i];
  best := winSum;
  for i := k to 5 do
  begin winSum := winSum - a[i-k] + a[i]; if winSum > best then best := winSum; end;
  WriteLn(best);
end."#
        ),
        &["20"]
    );
}
