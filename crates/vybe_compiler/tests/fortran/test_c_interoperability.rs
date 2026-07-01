use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(c_use_iso_01,"program p
use iso_c_binding
end program p
");
c!(c_ptr_02,"program p
use iso_c_binding
type(c_ptr) :: p
end program p
");
c!(c_funptr_03,"program p
use iso_c_binding
type(c_funptr) :: fp
end program p
");
c!(c_loc_04,"program p
use iso_c_binding
integer, target :: x
print *, c_associated(c_loc(x))
end program p
");
c!(c_f_pointer_05,"program p
use iso_c_binding
type(c_ptr) :: p
integer, pointer :: fp
call c_f_pointer(p, fp)
end program p
");
c!(c_funloc_06,"program p
use iso_c_binding
print *, c_associated(c_funloc(s))
contains
subroutine s() bind(c)
end subroutine s
end program p
");
c!(c_bind_type_07,"use iso_c_binding
type, bind(c) :: t
 integer(c_int) :: x
end type t
");
c!(c_string_08,"program p
use iso_c_binding
character(kind=c_char,len=4) :: s
print *, s
end program p
");
c!(c_array_09,"program p
use iso_c_binding
integer(c_int) :: a(3)
print *, a
end program p
");
c!(c_struct_10,"use iso_c_binding
type, bind(c) :: point
 integer(c_int) :: x
 integer(c_int) :: y
end type point
");
c!(c_bind_sub_11,"subroutine s() bind(c)
use iso_c_binding
end subroutine s
");
c!(c_bind_name_12,"subroutine s() bind(c,name='s_c')
use iso_c_binding
end subroutine s
");
c!(c_int_13,"program p
use iso_c_binding
integer(c_int) :: x
print *, x
end program p
");
c!(c_double_14,"program p
use iso_c_binding
real(c_double) :: x
print *, x
end program p
");
c!(c_bool_15,"program p
use iso_c_binding
logical(c_bool) :: x
print *, x
end program p
");
c!(c_size_t_16,"program p
use iso_c_binding
integer(c_size_t) :: x
print *, x
end program p
");
c!(c_char_17,"program p
use iso_c_binding
character(kind=c_char) :: x
print *, x
end program p
");
c!(c_null_ptr_18,"program p
use iso_c_binding
type(c_ptr) :: p
p = c_null_ptr
end program p
");
c!(c_null_funptr_19,"program p
use iso_c_binding
type(c_funptr) :: fp
fp = c_null_funptr
end program p
");
c!(c_associated_20,"program p
use iso_c_binding
type(c_ptr) :: p
print *, c_associated(p)
end program p
");