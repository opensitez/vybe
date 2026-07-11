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
