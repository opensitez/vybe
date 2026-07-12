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
    explicit_interfaces_01,
    "interface
subroutine s()
end subroutine s
end interface
"
);
c!(
    explicit_interfaces_02,
    "interface
subroutine s(x)
integer::x
end subroutine s
end interface
"
);
c!(
    explicit_interfaces_03,
    "interface
real function f(x)
real::x
end function f
end interface
"
);
c!(
    explicit_interfaces_04,
    "interface
integer function f()
end function f
end interface
"
);
c!(
    explicit_interfaces_05,
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
    explicit_interfaces_06,
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
    explicit_interfaces_07,
    "interface
subroutine s(a)
real::a(:)
end subroutine s
end interface
"
);
c!(
    explicit_interfaces_08,
    "interface
subroutine s(a)
integer, optional :: a
end subroutine s
end interface
"
);
c!(
    explicit_interfaces_09,
    "interface
subroutine s(a)
integer, value :: a
end subroutine s
end interface
"
);
c!(
    explicit_interfaces_10,
    "interface
subroutine s(a)
integer, intent(in) :: a
end subroutine s
end interface
"
);
