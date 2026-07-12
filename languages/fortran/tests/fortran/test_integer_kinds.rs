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
    integer_kinds_01,
    "program p
integer(kind=1) :: x
print *, x
end program p
"
);
c!(
    integer_kinds_02,
    "program p
integer(kind=2) :: x
print *, x
end program p
"
);
c!(
    integer_kinds_03,
    "program p
integer(kind=4) :: x
print *, x
end program p
"
);
c!(
    integer_kinds_04,
    "program p
integer(kind=8) :: x
print *, x
end program p
"
);
c!(
    integer_kinds_05,
    "program p
print *, selected_int_kind(2)
end program p
"
);
c!(
    integer_kinds_06,
    "program p
print *, selected_int_kind(9)
end program p
"
);
c!(
    integer_kinds_07,
    "program p
integer(kind=8) :: x=1_8
print *, x
end program p
"
);
c!(
    integer_kinds_08,
    "program p
integer(kind=4), parameter :: x=1_4
print *, x
end program p
"
);
c!(
    integer_kinds_09,
    "program p
integer(kind=8) :: a=1_8,b=2_8
print *, a+b
end program p
"
);
c!(
    integer_kinds_10,
    "program p
integer(kind=4) :: a=7_4
print *, mod(a,3_4)
end program p
"
);
