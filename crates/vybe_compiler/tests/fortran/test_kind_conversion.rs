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
    kind_conversion_01,
    "program p
integer :: i
real :: r=1.5
i = int(r)
print *, i
end program p
"
);
c!(
    kind_conversion_02,
    "program p
real :: r
r = real(1)
print *, r
end program p
"
);
c!(
    kind_conversion_03,
    "program p
double precision :: d
d = dble(1.0)
print *, d
end program p
"
);
c!(
    kind_conversion_04,
    "program p
complex :: z
z = cmplx(1.0,2.0)
print *, z
end program p
"
);
c!(
    kind_conversion_05,
    "program p
integer(kind=8) :: i
i = int(1.5, kind=8)
print *, i
end program p
"
);
c!(
    kind_conversion_06,
    "program p
real(kind=8) :: r
r = real(1, kind=8)
print *, r
end program p
"
);
c!(
    kind_conversion_07,
    "program p
complex(kind=8) :: z
z = cmplx(1.0_8,2.0_8,kind=8)
print *, z
end program p
"
);
c!(
    kind_conversion_08,
    "program p
integer :: i
i = nint(1.6)
print *, i
end program p
"
);
c!(
    kind_conversion_09,
    "program p
real :: r
r = transfer(1, r)
print *, r
end program p
"
);
c!(
    kind_conversion_10,
    "program p
integer :: i
real :: r=1.0
i = transfer(r, i)
print *, i
end program p
"
);
