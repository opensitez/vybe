/// Tests for sorting algorithms and search patterns in Pascal/Delphi
/// that go beyond test_programs_extended.rs (which already has bubble sort,
/// selection sort, binary search, linear search).

use super::helpers::run_pascal;

// ===================================================================
// INSERTION SORT
// ===================================================================

#[test] fn insertion_sort_basic() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    i, j, key: Integer;
begin
  SetLength(arr, 5);
  arr[0] := 5; arr[1] := 2; arr[2] := 4; arr[3] := 1; arr[4] := 3;
  for i := 1 to 4 do
  begin
    key := arr[i];
    j := i - 1;
    while (j >= 0) and (arr[j] > key) do
    begin
      arr[j + 1] := arr[j];
      j := j - 1;
    end;
    arr[j + 1] := key;
  end;
  for i := 0 to 4 do
    Write(IntToStr(arr[i]) + ' ');
  WriteLn('');
end."#), &["1 2 3 4 5 "]);
}

#[test] fn insertion_sort_strings() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of String;
    i, j: Integer;
    key: String;
begin
  SetLength(arr, 4);
  arr[0] := 'delta'; arr[1] := 'alpha'; arr[2] := 'charlie'; arr[3] := 'bravo';
  for i := 1 to 3 do
  begin
    key := arr[i];
    j := i - 1;
    while (j >= 0) and (arr[j] > key) do
    begin
      arr[j + 1] := arr[j];
      j := j - 1;
    end;
    arr[j + 1] := key;
  end;
  for i := 0 to 3 do
    WriteLn(arr[i]);
end."#), &["alpha", "bravo", "charlie", "delta"]);
}

// ===================================================================
// SHELL SORT
// ===================================================================

#[test] fn shell_sort_basic() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    n, gap, i, j, tmp: Integer;
begin
  SetLength(arr, 6);
  arr[0] := 64; arr[1] := 34; arr[2] := 25; arr[3] := 12; arr[4] := 22; arr[5] := 11;
  n := 6;
  gap := n div 2;
  while gap > 0 do
  begin
    for i := gap to n - 1 do
    begin
      tmp := arr[i];
      j := i;
      while (j >= gap) and (arr[j - gap] > tmp) do
      begin
        arr[j] := arr[j - gap];
        j := j - gap;
      end;
      arr[j] := tmp;
    end;
    gap := gap div 2;
  end;
  for i := 0 to n - 1 do
    Write(IntToStr(arr[i]) + ' ');
  WriteLn('');
end."#), &["11 12 22 25 34 64 "]);
}

// ===================================================================
// COUNTING SORT
// ===================================================================

#[test] fn counting_sort_small_range() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    counts: array[0..9] of Integer;
    i, j, idx: Integer;
begin
  SetLength(arr, 7);
  arr[0] := 4; arr[1] := 2; arr[2] := 2; arr[3] := 8;
  arr[4] := 3; arr[5] := 3; arr[6] := 1;
  for i := 0 to 9 do counts[i] := 0;
  for i := 0 to 6 do Inc(counts[arr[i]]);
  idx := 0;
  for i := 0 to 9 do
    for j := 1 to counts[i] do
    begin
      arr[idx] := i;
      Inc(idx);
    end;
  for i := 0 to 6 do
    Write(IntToStr(arr[i]) + ' ');
  WriteLn('');
end."#), &["1 2 2 3 3 4 8 "]);
}

// ===================================================================
// ARRAY MINIMUM AND MAXIMUM
// ===================================================================

#[test] fn find_min_max_in_array() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    mn, mx, i: Integer;
begin
  SetLength(arr, 6);
  arr[0] := 3; arr[1] := 7; arr[2] := 1; arr[3] := 9; arr[4] := 4; arr[5] := 6;
  mn := arr[0];
  mx := arr[0];
  for i := 1 to 5 do
  begin
    if arr[i] < mn then mn := arr[i];
    if arr[i] > mx then mx := arr[i];
  end;
  WriteLn(mn);
  WriteLn(mx);
end."#), &["1", "9"]);
}

#[test] fn find_second_largest() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    first, second, i: Integer;
begin
  SetLength(arr, 5);
  arr[0] := 5; arr[1] := 2; arr[2] := 9; arr[3] := 7; arr[4] := 3;
  first := -1;
  second := -1;
  for i := 0 to 4 do
  begin
    if arr[i] > first then
    begin
      second := first;
      first := arr[i];
    end
    else if arr[i] > second then
      second := arr[i];
  end;
  WriteLn(second);
end."#), &["7"]);
}

// ===================================================================
// PRIME SIEVE (SIEVE OF ERATOSTHENES)
// ===================================================================

#[test] fn prime_sieve_count() {
    assert_eq!(run_pascal(r#"program T;
var sieve: array[2..30] of Boolean;
    i, j, count: Integer;
begin
  for i := 2 to 30 do sieve[i] := True;
  i := 2;
  while i * i <= 30 do
  begin
    if sieve[i] then
    begin
      j := i * i;
      while j <= 30 do
      begin
        sieve[j] := False;
        j := j + i;
      end;
    end;
    Inc(i);
  end;
  count := 0;
  for i := 2 to 30 do
    if sieve[i] then Inc(count);
  WriteLn(count);
end."#), &["10"]);
}

#[test] fn prime_sieve_list() {
    assert_eq!(run_pascal(r#"program T;
var sieve: array[2..20] of Boolean;
    i, j: Integer;
begin
  for i := 2 to 20 do sieve[i] := True;
  i := 2;
  while i * i <= 20 do
  begin
    if sieve[i] then
    begin
      j := i * i;
      while j <= 20 do
      begin
        sieve[j] := False;
        j := j + i;
      end;
    end;
    Inc(i);
  end;
  for i := 2 to 20 do
    if sieve[i] then Write(IntToStr(i) + ' ');
  WriteLn('');
end."#), &["2 3 5 7 11 13 17 19 "]);
}

// ===================================================================
// TWO POINTER TECHNIQUE
// ===================================================================

#[test] fn two_sum_sorted_array() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    lo, hi, target: Integer;
    found: Boolean;
begin
  SetLength(arr, 6);
  arr[0] := 1; arr[1] := 2; arr[2] := 3; arr[3] := 5; arr[4] := 7; arr[5] := 9;
  target := 10;
  lo := 0;
  hi := 5;
  found := False;
  while lo < hi do
  begin
    if arr[lo] + arr[hi] = target then
    begin
      WriteLn(arr[lo]);
      WriteLn(arr[hi]);
      found := True;
      Break;
    end
    else if arr[lo] + arr[hi] < target then
      Inc(lo)
    else
      Dec(hi);
  end;
  if not found then WriteLn('not found');
end."#), &["3", "7"]);
}

// ===================================================================
// FREQUENCY COUNT
// ===================================================================

#[test] fn frequency_of_elements() {
    assert_eq!(run_pascal(r#"program T;
var data: array of Integer;
    freq: array[1..5] of Integer;
    i: Integer;
begin
  SetLength(data, 8);
  data[0] := 1; data[1] := 2; data[2] := 3; data[3] := 2;
  data[4] := 1; data[5] := 2; data[6] := 5; data[7] := 3;
  for i := 1 to 5 do freq[i] := 0;
  for i := 0 to 7 do
    if (data[i] >= 1) and (data[i] <= 5) then
      Inc(freq[data[i]]);
  for i := 1 to 5 do
    WriteLn(IntToStr(i) + ':' + IntToStr(freq[i]));
end."#), &["1:2", "2:3", "3:2", "4:0", "5:1"]);
}

// ===================================================================
// MERGE TWO SORTED ARRAYS
// ===================================================================

#[test] fn merge_sorted_arrays() {
    assert_eq!(run_pascal(r#"program T;
var a, b, c: array of Integer;
    i, j, k: Integer;
begin
  SetLength(a, 3); SetLength(b, 3); SetLength(c, 6);
  a[0] := 1; a[1] := 3; a[2] := 5;
  b[0] := 2; b[1] := 4; b[2] := 6;
  i := 0; j := 0; k := 0;
  while (i < 3) and (j < 3) do
  begin
    if a[i] <= b[j] then
    begin
      c[k] := a[i]; Inc(i);
    end
    else
    begin
      c[k] := b[j]; Inc(j);
    end;
    Inc(k);
  end;
  while i < 3 do begin c[k] := a[i]; Inc(i); Inc(k); end;
  while j < 3 do begin c[k] := b[j]; Inc(j); Inc(k); end;
  for i := 0 to 5 do
    Write(IntToStr(c[i]) + ' ');
  WriteLn('');
end."#), &["1 2 3 4 5 6 "]);
}

// ===================================================================
// ARRAY ROTATION
// ===================================================================

#[test] fn rotate_array_left() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    first, i: Integer;
begin
  SetLength(arr, 5);
  arr[0] := 1; arr[1] := 2; arr[2] := 3; arr[3] := 4; arr[4] := 5;
  first := arr[0];
  for i := 0 to 3 do
    arr[i] := arr[i + 1];
  arr[4] := first;
  for i := 0 to 4 do
    Write(IntToStr(arr[i]) + ' ');
  WriteLn('');
end."#), &["2 3 4 5 1 "]);
}

// ===================================================================
// RUNNING SUM (PREFIX SUM)
// ===================================================================

#[test] fn prefix_sum_array() {
    assert_eq!(run_pascal(r#"program T;
var arr, prefix: array of Integer;
    i: Integer;
begin
  SetLength(arr, 5);
  SetLength(prefix, 5);
  arr[0] := 1; arr[1] := 2; arr[2] := 3; arr[3] := 4; arr[4] := 5;
  prefix[0] := arr[0];
  for i := 1 to 4 do
    prefix[i] := prefix[i - 1] + arr[i];
  for i := 0 to 4 do
    Write(IntToStr(prefix[i]) + ' ');
  WriteLn('');
end."#), &["1 3 6 10 15 "]);
}

// ===================================================================
// DUTCH FLAG PARTITION
// ===================================================================

#[test] fn partition_around_pivot() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    i, j, pivot, tmp: Integer;
begin
  SetLength(arr, 6);
  arr[0] := 3; arr[1] := 1; arr[2] := 4; arr[3] := 1; arr[4] := 5; arr[5] := 2;
  pivot := 3;
  i := 0;
  for j := 0 to 5 do
    if arr[j] <= pivot then
    begin
      tmp := arr[i]; arr[i] := arr[j]; arr[j] := tmp;
      Inc(i);
    end;
  WriteLn(arr[0] <= pivot);
  WriteLn(arr[5] >= pivot);
end."#), &["true", "true"]);
}

// ===================================================================
// LONGEST INCREASING SUBSEQUENCE LENGTH
// ===================================================================

#[test] fn lis_length() {
    assert_eq!(run_pascal(r#"program T;
var arr: array of Integer;
    dp: array of Integer;
    i, j, best: Integer;
begin
  SetLength(arr, 6);
  arr[0] := 3; arr[1] := 1; arr[2] := 4; arr[3] := 1; arr[4] := 5; arr[5] := 9;
  SetLength(dp, 6);
  for i := 0 to 5 do
  begin
    dp[i] := 1;
    for j := 0 to i - 1 do
      if (arr[j] < arr[i]) and (dp[j] + 1 > dp[i]) then
        dp[i] := dp[j] + 1;
  end;
  best := 0;
  for i := 0 to 5 do
    if dp[i] > best then best := dp[i];
  WriteLn(best);
end."#), &["4"]);
}

// ===================================================================
// MATRIX TRANSPOSE
// ===================================================================

#[test] fn matrix_transpose() {
    assert_eq!(run_pascal(r#"program T;
var mat: array[0..1, 0..2] of Integer;
    trans: array[0..2, 0..1] of Integer;
    i, j: Integer;
begin
  mat[0, 0] := 1; mat[0, 1] := 2; mat[0, 2] := 3;
  mat[1, 0] := 4; mat[1, 1] := 5; mat[1, 2] := 6;
  for i := 0 to 1 do
    for j := 0 to 2 do
      trans[j, i] := mat[i, j];
  for i := 0 to 2 do
    WriteLn(IntToStr(trans[i, 0]) + ' ' + IntToStr(trans[i, 1]));
end."#), &["1 4", "2 5", "3 6"]);
}
