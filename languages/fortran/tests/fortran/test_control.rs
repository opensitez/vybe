use super::helpers::{assert_fortran, compile_ok};

macro_rules! c {
    ($n:ident, $s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };

    ($n:ident, $s:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $n() {
            assert_fortran($s, &[$($expected),*]);
        }
    };
}

c!(
    ctrl_associate_01,
    "program p
integer :: x, y
x = 1
associate (a => x)
 a = a + 4
 y = a
end associate
print *, y
print *, x
end program p
",
    ["5", "5"]
);

c!(
    ctrl_block_02,
    "program p
block
 integer :: x
 x = 7
 print *, x
end block
end program p
",
    ["7"]
);

c!(
    ctrl_change_team_03,
    "program p
change team (team=1)
end team
end program p
"
);

c!(
    ctrl_critical_04,
    "program p
critical
 print *,1
end critical
end program p
"
);

c!(
    ctrl_error_stop_05,
    "program p
error stop
end program p
"
);

c!(
    ctrl_stop_code_06,
    "program p
stop 1
end program p
"
);

c!(
    ctrl_cycle_07,
    "program p
integer::i, sum
sum = 0
do i = 1, 5
    if (mod(i, 2) /= 0) cycle
    sum = sum + i
end do
print *, sum
end program p
",
    ["6"]
);

c!(
    ctrl_exit_08,
    "program p
integer::i, sum
sum = 0
do i = 1, 100
    if (i > 5) exit
    sum = sum + i
end do
print *, sum
end program p
",
    ["15"]
);

c!(
    ctrl_return_09,
    "program test
integer :: x
call s(x)
print *, x
contains
subroutine s(a)
    integer, intent(out) :: a
    a = 1
    return
    a = 2
end subroutine s
end program test
",
    ["1"]
);

c!(
    ctrl_goto_computed_10,
    "program p
integer :: k, x
k = 2
x = 99
go to (10,20,30), k
10 x = 10
goto 40
20 x = 20
goto 40
30 x = 30
40 continue
print *, x
end program p
",
    ["20"]
);

c!(
    ctrl_goto_assigned_11,
    "program p
integer :: n
integer :: x
assign 10 to n
x = 0
go to n
print *, 2
10 continue
print *, 10
end program p
",
    ["10"]
);

c!(
    ctrl_arith_if_12,
    "program p
integer :: x, y
x = 1
if (x) 10,20,30
10 y = 10
goto 40
20 y = 20
goto 40
30 y = 30
40 continue
print *, y
end program p
",
    ["10"]
);

c!(
    ctrl_arith_if_zero_12b,
    "program p
integer :: x, y
x = 0
if (x) 10,20,30
10 y = 10
goto 40
20 y = 20
goto 40
30 y = 30
40 continue
print *, y
end program p
",
    ["20"]
);

c!(
    ctrl_arith_if_negative_12a,
    "program p
integer :: x, y
x = -1
if (x) 10,20,30
10 y = 10
goto 40
20 y = 20
goto 40
30 y = 30
40 continue
print *, y
end program p
",
    ["10"]
);

c!(
    ctrl_block_if_13,
    "program p
integer::x=1
if (x==1) then
 print *,1
else
 print *,2
end if
end program p
",
    ["1"]
);

c!(
    ctrl_select_case_14,
    "program p
integer::x=2
select case(x)
 case(1)
  print *,1
 case(2)
  print *,2
 case default
  print *,3
end select
end program p
",
    ["2"]
);

c!(
    ctrl_do_15,
    "program p
integer::i, sum
sum = 0
do i = 1, 4
 sum = sum + i
end do
print *, sum
end program p
",
    ["10"]
);

c!(
    ctrl_do_while_16,
    "program p
integer::i
i = 0
do while(i<3)
 i = i + 1
end do
print *, i
end program p
",
    ["3"]
);

c!(
    ctrl_do_concurrent_17,
    "program p
integer::i
do concurrent (i=1:3)
 print *,i
end do
end program p
",
    ["1", "2", "3"]
);

c!(
    ctrl_do_concurrent_fills_array,
    "program p
integer::i
integer::a(3)
do concurrent (i=1:3)
  a(i) = i + 5
end do
print *, a(1)
print *, a(2)
print *, a(3)
end program p
",
    ["6", "7", "8"]
);

c!(
    ctrl_where_18,
    "program p
integer::a(3)=[1,2,3]
where(a>1) a=a+1
print *, a(1)
print *, a(2)
print *, a(3)
end program p
",
    ["1", "3", "4"]
);

c!(
    ctrl_forall_19,
    "program p
integer::a(3), i
forall(i=1:3) a(i)=i*2
print *, a(1)
print *, a(2)
print *, a(3)
end program p
",
    ["2", "4", "6"]
);

c!(
    ctrl_select_type_20,
    "program p
class(*),allocatable::x
allocate(integer::x)
x = 1
select type(x)
 type is(integer)
  print *,x
 class default
  print *, 2
end select
end program p
",
    ["1"]
);

c!(
    ctrl_exit_nested_21,
    "program p
integer::outer_i, inner_i, total
total = 0
outer: do outer_i = 1, 4
  do inner_i = 1, 3
    if (outer_i == 3 .and. inner_i == 2) exit outer
    total = total + 1
  end do
end do outer
print *, total
end program p
",
    ["7"]
);

c!(
    ctrl_cycle_named_loop_22,
    "program p
integer::i, j, total
total = 0
row: do i = 1, 2
  col: do j = 1, 5
    if (mod(j, 2) == 0) cycle row
    total = total + 1
  end do col
end do row
print *, total
end program p
",
    ["2"]
);

c!(
    ctrl_do_descending_23,
    "program p
integer::i, total
total = 0
do i = 5, 1, -2
  total = total + i
end do
print *, total
end program p
",
    ["9"]
);

c!(
    ctrl_where_elsewhere_24,
    "program p
integer::a(4)
integer::b(4)
a = (/1, 2, 3, 4/)
b = 0
where (a <= 2)
  b = 10
elsewhere (a <= 3)
  b = 20
elsewhere
  b = 30
end where
print *, b(1)
print *, b(2)
print *, b(3)
print *, b(4)
end program p
",
    ["10", "10", "20", "30"]
);

c!(
    ctrl_select_type_25,
    "program p
class(*), allocatable :: x
allocate(real :: x)
select type (x)
  type is (integer)
    print *, 1
  type is (real)
    print *, 2
  class default
    print *, 3
end select
end program p
",
    ["2"]
);

c!(
    ctrl_stop_then_return_26,
    "program p\ninteger :: i\nprint *, 'before'\ni = 1\nif (i == 1) then\n    return\nend if\nprint *, i\nend program p\n",
    ["before"]
);

c!(
    ctrl_return_27,
    "program p\ninteger :: i\ninteger :: s\nprint *, 'start'\nif (.true.) then\n    i = 1\n    s = 2\n    return\nend if\nprint *, i\nprint *, s\nend program p\n",
    ["start"]
);

c!(
    ctrl_do_named_while_named_cycle_28,
    "program p\ninteger :: i, c\ni = 0\nc = 0\nspin: do while (i < 6)\n    i = i + 1\n    if (mod(i, 2) == 0) cycle spin\n    c = c + 1\nend do spin\nprint *, c\nend program p\n",
    ["3"]
);

c!(
    ctrl_do_named_while_exit_29,
    "program p\ninteger :: i, s\ni = 0\ns = 0\nspin: do while (i < 10)\n    i = i + 1\n    if (i == 4) exit spin\n    s = s + 1\nend do spin\nprint *, s\nend program p\n",
    ["3"]
);
