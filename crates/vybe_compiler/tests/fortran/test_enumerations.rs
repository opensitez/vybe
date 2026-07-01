use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(enum_bindc_01,"enum, bind(c)
enumerator :: red
end enum
");
c!(enum_two_02,"enum, bind(c)
enumerator :: red=1, green=2
end enum
");
c!(enum_three_03,"enum, bind(c)
enumerator :: a=1, b=2, c=3
end enum
");
c!(enum_assign_04,"enum, bind(c)
enumerator :: red=1
end enum
program p
integer :: x
x = red
print *, x
end program p
");
c!(enum_compare_05,"enum, bind(c)
enumerator :: red=1, blue=2
end enum
program p
logical :: l
l = red < blue
print *, l
end program p
");
c!(enum_default_06,"enum, bind(c)
enumerator :: a, b, c
end enum
");
c!(enum_negative_07,"enum, bind(c)
enumerator :: a=-1, b=0
end enum
");
c!(enum_large_08,"enum, bind(c)
enumerator :: big=1000
end enum
");
c!(enum_use_in_select_09,"enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
select case(a)
 case(1)
  print *,1
end select
end program p
");
c!(enum_with_module_10,"module m
enum, bind(c)
enumerator :: a=1
end enum
end module m
");
c!(enum_reuse_11,"module m
enum, bind(c)
enumerator :: first=1, second=2
end enum
contains
subroutine s()
print *, first
end subroutine s
end module m
");
c!(enum_arith_12,"enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
print *, a+b
end program p
");
c!(enum_if_13,"enum, bind(c)
enumerator :: a=1
end enum
program p
if (a == 1) print *,1
end program p
");
c!(enum_case_14,"enum, bind(c)
enumerator :: a=1
end enum
program p
select case(a)
case (1)
 print *,1
end select
end program p
");
c!(enum_print_15,"enum, bind(c)
enumerator :: a=1
end enum
program p
print *, a
end program p
");
c!(enum_param_like_16,"enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
integer, parameter :: x = a
print *, x
end program p
");
c!(enum_expr_17,"enum, bind(c)
enumerator :: a=1, b=2
end enum
program p
integer :: x
x = a * b
print *, x
end program p
");
c!(enum_module_use_18,"module m
enum, bind(c)
enumerator :: a=1
end enum
end module m
program p
use m
print *, a
end program p
");
c!(enum_kind_int_19,"enum, bind(c)
enumerator :: a=1
end enum
program p
integer :: x
x = a
end program p
");
c!(enum_named_values_20,"enum, bind(c)
enumerator :: sunday=0, monday=1
end enum
");