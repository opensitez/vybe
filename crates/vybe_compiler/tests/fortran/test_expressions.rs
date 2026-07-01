use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(expr_add_01, "program p
integer :: a=1,b=2,c
c = a + b
print *, c
end program p
");
c!(expr_sub_02, "program p
integer :: a=5,b=2,c
c = a - b
print *, c
end program p
");
c!(expr_mul_03, "program p
integer :: a=3,b=4,c
c = a * b
print *, c
end program p
");
c!(expr_div_04, "program p
real :: a=8.0,b=2.0,c
c = a / b
print *, c
end program p
");
c!(expr_pow_05, "program p
integer :: a=2,b
a = 2
b = a ** 3
print *, b
end program p
");
c!(expr_unary_06, "program p
integer :: a
a = -5
print *, a
end program p
");
c!(expr_paren_07, "program p
integer :: x
x = (2 + 3) * 4
print *, x
end program p
");
c!(expr_prec_08, "program p
integer :: x
x = 2 + 3 * 4
print *, x
end program p
");
c!(expr_logical_and_09, "program p
logical :: x
x = .true. .and. .false.
print *, x
end program p
");
c!(expr_logical_or_10, "program p
logical :: x
x = .true. .or. .false.
print *, x
end program p
");
c!(expr_logical_not_11, "program p
logical :: x
x = .not. .false.
print *, x
end program p
");
c!(expr_eq_12, "program p
logical :: x
x = 1 == 1
print *, x
end program p
");
c!(expr_ne_13, "program p
logical :: x
x = 1 /= 2
print *, x
end program p
");
c!(expr_lt_14, "program p
logical :: x
x = 1 < 2
print *, x
end program p
");
c!(expr_le_15, "program p
logical :: x
x = 1 <= 2
print *, x
end program p
");
c!(expr_gt_16, "program p
logical :: x
x = 2 > 1
print *, x
end program p
");
c!(expr_ge_17, "program p
logical :: x
x = 2 >= 1
print *, x
end program p
");
c!(expr_concat_18, "program p
character(len=2) :: s
s = 'a'//'b'
print *, s
end program p
");
c!(expr_char_rel_19, "program p
logical :: x
x = 'a' < 'b'
print *, x
end program p
");
c!(expr_complex_add_20, "program p
complex :: a=(1.0,2.0), b=(3.0,4.0), c
c = a + b
print *, c
end program p
");
c!(expr_array_constructor_21, "program p
integer :: a(3)
a = [1,2,3]
print *, a
end program p
");
c!(expr_section_22, "program p
integer :: a(4)
a = [1,2,3,4]
print *, a(2:3)
end program p
");
c!(expr_index_23, "program p
integer :: a(3)
a = [1,2,3]
print *, a(2)
end program p
");
c!(expr_func_call_24, "program p
print *, abs(-3)
end program p
");
c!(expr_nested_call_25, "program p
print *, max(1, min(2,3))
end program p
");
c!(expr_kind_conv_26, "program p
integer :: i
real :: r=1.5
i = int(r)
print *, i
end program p
");
c!(expr_real_conv_27, "program p
real :: r
r = real(3)
print *, r
end program p
");
c!(expr_merge_28, "program p
integer :: x
x = merge(1,2,.true.)
print *, x
end program p
");
c!(expr_implied_do_29, "program p
integer :: a(3)
a = [(i, i=1,3)]
print *, a
end program p
");
c!(expr_masked_where_30, "program p
integer :: a(3)=[1,2,3]
where (a > 1) a = a + 1
print *, a
end program p
");