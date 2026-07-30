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
    implicit_interfaces_01,
    "external s
call s()
end
"
);
c!(
    implicit_interfaces_02,
    "integer f
external f
print *, f()
end
"
);
c!(
    implicit_interfaces_03,
    "program p
external s
call s(1)
end program p
"
);
c!(
    implicit_interfaces_04,
    "program p
external s
integer :: x
x=1
call s(x)
end program p
"
);
c!(
    implicit_interfaces_05,
    "subroutine caller()
external s
call s()
end subroutine caller
"
);
c!(
    implicit_interfaces_06,
    "program p
external f
real :: x
x = f()
end program p
"
);
c!(
    implicit_interfaces_07,
    "program p
external s
call s('a')
end program p
"
);
c!(
    implicit_interfaces_08,
    "program p
external s
call s(1.0)
end program p
"
);
c!(
    implicit_interfaces_09,
    "program p
external s
call s(.true.)
end program p
"
);
c!(
    implicit_interfaces_10,
    "program p
external s
call s((1.0,2.0))
end program p
"
);
c!(
    implicit_interfaces_11,
    "program p
external f
integer :: x
x = f()
end program p
"
);
c!(
    implicit_interfaces_12,
    "program p
external f
real :: x, y
x = f(y, 2)
end program p
"
);
c!(
    implicit_interfaces_13,
    "program p
external f
logical :: ok
ok = f(2, 3)
end program p
"
);
c!(
    implicit_interfaces_14,
    "program p
external f
complex :: z
z = f(1.0, 2.0)
end program p
"
);
c!(
    implicit_interfaces_15,
    "program p
external f
character(len=10) :: text
text = f()
end program p
"
);
c!(
    implicit_interfaces_16,
    "program p
integer, dimension(3) :: arr
external s
call s(arr)
end program p
"
);
c!(
    implicit_interfaces_17,
    "program p
integer :: i, j
external mix
call mix(i, j, i + j)
end program p
"
);
c!(
    implicit_interfaces_18,
    "subroutine outer()
external s
call s(1)
end subroutine outer
"
);
c!(
    implicit_interfaces_19,
    "program p
implicit none
call caller()
contains
subroutine caller()
integer :: x
external f
x = f()
end subroutine caller
end program p
"
);
c!(
    implicit_interfaces_20,
    "program p
implicit none
integer :: x
external f
x = f(x)
end program p
"
);
