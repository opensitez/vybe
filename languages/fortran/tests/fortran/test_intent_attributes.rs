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
    intent_attributes_01,
    "subroutine s(x)
integer, intent(in) :: x
end subroutine s
"
);
c!(
    intent_attributes_02,
    "subroutine s(x)
integer, intent(out) :: x
end subroutine s
"
);
c!(
    intent_attributes_03,
    "subroutine s(x)
integer, intent(inout) :: x
end subroutine s
"
);
c!(
    intent_attributes_04,
    "subroutine s(a)
real, intent(in) :: a(:)
end subroutine s
"
);
c!(
    intent_attributes_05,
    "subroutine s(a)
real, intent(out) :: a(:)
end subroutine s
"
);
c!(
    intent_attributes_06,
    "subroutine s(a)
real, intent(inout) :: a(:)
end subroutine s
"
);
c!(
    intent_attributes_07,
    "subroutine s(x)
character(len=*), intent(in) :: x
end subroutine s
"
);
c!(
    intent_attributes_08,
    "subroutine s(x)
logical, intent(out) :: x
end subroutine s
"
);
c!(
    intent_attributes_09,
    "subroutine s(x)
complex, intent(inout) :: x
end subroutine s
"
);
c!(
    intent_attributes_10,
    "subroutine s(x)
integer, optional, intent(in) :: x
end subroutine s
"
);
