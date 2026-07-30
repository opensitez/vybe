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
    generic_ambiguity_01,
    "module m
interface g
module procedure si,sr
end interface
contains
subroutine si(i)
integer::i
end
subroutine sr(r)
real::r
end
end module m
"
);
c!(
    generic_ambiguity_02,
    "module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(j)
integer::j
end
end module m
"
);
c!(
    generic_ambiguity_03,
    "module m
interface g
module procedure s1,s2,s3
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
subroutine s3(c)
complex::c
end
end module m
"
);
c!(
    generic_ambiguity_04,
    "module m
interface operator(+)
module procedure addi,addr
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
real function addr(a,b)
real::a,b
addr=a+b
end
end module m
"
);
c!(
    generic_ambiguity_05,
    "module m
interface assignment(=)
module procedure asgi,asgr
end interface
contains
subroutine asgi(a,b)
integer::a,b
a=b
end
subroutine asgr(a,b)
real::a,b
a=b
end
end module m
"
);
c!(
    generic_ambiguity_06,
    "module m
interface g
module procedure s1
end interface
contains
subroutine s1(i)
integer::i
end
end module m
program p
use m
call g(1)
end program p
"
);
c!(
    generic_ambiguity_07,
    "module m
interface g
module procedure s1
end interface
contains
subroutine s1(r)
real::r
end
end module m
program p
use m
call g(1.0)
end program p
"
);
c!(
    generic_ambiguity_08,
    "module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(i)
integer::i
end
subroutine s2(r)
real::r
end
end module m
program p
use m
call g(1)
call g(1.0)
end program p
"
);
c!(
    generic_ambiguity_09,
    "module m
interface operator(-)
module procedure subi,subr
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
real function subr(a,b)
real::a,b
subr=a-b
end
end module m
"
);
c!(
    generic_ambiguity_10,
    "module m
interface g
module procedure s1,s2
end interface
contains
subroutine s1(c)
character(len=*)::c
end
subroutine s2(l)
logical::l
end
end module m
"
);

#[test]
fn generic_ambiguity_function_call_runtime_dispatch() {
    let out = run_prints(
        r#"
module m
    interface g
        module procedure gi, gr
    end interface
contains
    integer function gi(i)
        integer, intent(in) :: i
        gi = i + 10
    end function

    real function gr(r)
        real, intent(in) :: r
        gr = r + 1.5
    end function
end module m

program test_generic_ambiguity_function_call_runtime_dispatch
    use m
    print *, g(4)
    print *, nint(g(2.0))
end program test_generic_ambiguity_function_call_runtime_dispatch
"#,
    );
    assert_eq!(out, vec!["14", "4"]);
}

#[test]
fn generic_ambiguity_subroutine_call_runtime_dispatch() {
    let out = run_prints(
        r#"
module m
    interface g
        module procedure si, sr
    end interface
contains
    subroutine si(i)
        integer, intent(in) :: i
        print *, i + 10
    end subroutine

    subroutine sr(r)
        real, intent(in) :: r
        print *, nint(r) + 20
    end subroutine
end module m

program test_generic_ambiguity_subroutine_call_runtime_dispatch
    use m
    call g(2)
    call g(3.0)
end program test_generic_ambiguity_subroutine_call_runtime_dispatch
"#,
    );
    assert_eq!(out, vec!["12", "23"]);
}

#[test]
fn generic_ambiguity_function_like_dispatch_with_result_types() {
    let out = run_prints(
        r#"
module m
    interface g
        module procedure i2, r2
    end interface
contains
    integer function i2(i)
        integer, intent(in) :: i
        i2 = i * 2
    end function

    real function r2(r)
        real, intent(in) :: r
        r2 = r * 3.0
    end function
end module m

program test_generic_ambiguity_function_like_dispatch_with_result_types
    use m
    print *, g(3)
    print *, nint(g(1.5))
end program test_generic_ambiguity_function_like_dispatch_with_result_types
"#,
    );
    assert_eq!(out, vec!["6", "5"]);
}
