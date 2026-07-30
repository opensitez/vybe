use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    interface_blocks_01,
    "interface
subroutine s()
end subroutine s
end interface
"
);
c!(
    interface_blocks_02,
    "interface
subroutine s(x)
integer::x
end subroutine s
end interface
"
);
c!(
    interface_blocks_03,
    "interface
real function f(x)
real::x
end function f
end interface
"
);
c!(
    interface_blocks_04,
    "module m
interface
subroutine s(x)
integer::x
end subroutine s
end interface
end module m
"
);
c!(
    interface_blocks_05,
    "program p
interface
subroutine s(x)
integer::x
end subroutine s
end interface
call s(1)
end program p
"
);
c!(
    interface_blocks_06,
    "interface
integer function f()
end function f
end interface
"
);
c!(
    interface_blocks_07,
    "interface
subroutine s(a)
real::a(:)
end subroutine s
end interface
"
);
c!(
    interface_blocks_08,
    "interface
subroutine s(x)
integer, optional :: x
end subroutine s
end interface
"
);
c!(
    interface_blocks_09,
    "interface
subroutine s(x)
integer, value :: x
end subroutine s
end interface
"
);
c!(
    interface_blocks_10,
    "interface
subroutine s(x)
integer, intent(in) :: x
end subroutine s
end interface
"
);

c!(
    interface_blocks_11,
    "interface
subroutine s(x, y)
integer, intent(inout) :: x
integer, intent(in) :: y
end subroutine s
end interface
"
);

c!(
    interface_blocks_12,
    "interface
integer function f(a, b, scale)
integer, intent(in) :: a, b
integer, intent(in), optional :: scale
end function f
end interface
"
);

c!(
    interface_blocks_13,
    "interface
subroutine copy(src, dst)
integer, intent(in) :: src(:)
integer, intent(out) :: dst(:)
end subroutine copy
end interface
"
);

c!(
    interface_blocks_14,
    "interface
logical function has_value(v)
integer, intent(in) :: v
end function has_value
end interface
"
);

c!(
    interface_blocks_15,
    "interface
character(len=4) function aschar(i)
integer, intent(in) :: i
end function aschar
end interface
"
);

c!(
    interface_blocks_16,
    "interface operator(.custom.)
integer function custom_add(a, b)
integer, intent(in) :: a, b
end function custom_add
end interface
"
);

c!(
    interface_blocks_17,
    "interface assignment(=)
subroutine assign_wrapper(lhs, rhs)
integer, intent(out) :: lhs
integer, intent(in) :: rhs
end subroutine assign_wrapper
end interface\n"
);

c!(
    interface_blocks_18,
    "module m
implicit none\n
interface add_one
module procedure add_one_impl
end interface
contains
function add_one_impl(x) result(r)
integer, intent(in) :: x\ninteger :: r\nr = x + 1\nend function add_one_impl
end module m
"
);

c!(
    interface_blocks_19,
    "interface
subroutine f(x)
integer, pointer :: x
end subroutine f
end interface
"
);

c!(
    interface_blocks_20,
    "interface
subroutine s(x)\ninteger, optional, intent(in) :: x\nend subroutine s\nend interface\n"
);
