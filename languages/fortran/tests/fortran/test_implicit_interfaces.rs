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
