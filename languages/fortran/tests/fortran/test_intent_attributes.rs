use super::helpers::{compile_ok, run_prints};
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

c!(
    intent_attributes_11,
    "subroutine s(x, y)
integer, intent(in) :: x
integer, optional, intent(in) :: y
end subroutine s
"
);

c!(
    intent_attributes_12,
    "subroutine s(x)
integer, intent(in) :: x
integer :: y
y = x + 1
end subroutine s
"
);

c!(
    intent_attributes_13,
    "subroutine s(a, b, c)
integer, intent(inout) :: a
integer, intent(in) :: b
integer, intent(out) :: c
c = a + b
a = c
end subroutine s
"
);

c!(
    intent_attributes_14,
    "program p
integer :: i
call s(i)
contains
subroutine s(a)
integer, intent(inout) :: a
a = a + 1
end subroutine s
end program p
"
);

#[test]
fn intent_attributes_runtime_inout_update() {
    let out = run_prints(
        r#"
program test_intent_attributes
integer :: x = 2
call bump(x)
print *, x

contains
subroutine bump(v)
integer, intent(inout) :: v
v = v + 3
end subroutine bump
end program test_intent_attributes
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn intent_attributes_runtime_out_initialization() {
    let out = run_prints(
        r#"
program test_intent_attributes
integer :: x
integer :: y = 2
call write_out(y, x)
print *, x

contains
subroutine write_out(src, dst)
integer, intent(in) :: src
integer, intent(out) :: dst
dst = src * 2
end subroutine write_out
end program test_intent_attributes
"#,
    );

    assert_eq!(out, vec!["4"]);
}

#[test]
fn intent_attributes_runtime_optional_absent_kept_defaulted() {
    let out = run_prints(
        r#"
program test_intent_attributes
integer :: x = 1
call log_value(x)
call log_value(2)
contains
subroutine log_value(x, scale)
integer, intent(in) :: x
integer, optional, intent(in) :: scale
if (present(scale)) then
print *, x * scale
else
print *, x
end if
end subroutine log_value
end program test_intent_attributes
"#,
    );

    assert_eq!(out, vec!["1", "4"]);
}

#[test]
fn intent_attributes_runtime_array_inout_section_update() {
    let out = run_prints(
        r#"
program test_intent_attributes
integer :: a(3) = [1,2,3]
call inc_section(a(1:3:2))
print *, a(1)
print *, a(3)

contains
subroutine inc_section(x)
integer, intent(inout) :: x(:)
x = x + 1
end subroutine inc_section
end program test_intent_attributes
"#,
    );

    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn intent_attributes_runtime_character_inout() {
    let out = run_prints(
        r#"
program test_intent_attributes
character(len=8) :: word = "a"
call append_char(word)
print *, word

contains
subroutine append_char(s)
character(len=*), intent(inout) :: s
s = trim(s) // "bc"
end subroutine append_char
end program test_intent_attributes
"#,
    );

    assert_eq!(out, vec!["abc"]);
}
