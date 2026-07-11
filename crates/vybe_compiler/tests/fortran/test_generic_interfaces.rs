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
    gen_if_01,
    "module m
interface g
module procedure s1
end interface
contains
subroutine s1()
end subroutine s1
end module m
"
);
c!(
    gen_if_02,
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
    gen_if_03,
    "module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
end module m
"
);
c!(
    gen_if_04,
    "module m
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a,b)
integer::a,b
a=b
end
end module m
"
);
c!(
    gen_if_05,
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
    gen_if_06,
    "module m
interface operator(-)
module procedure subi
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
end module m
"
);
c!(
    gen_if_07,
    "module m
interface operator(*)
module procedure muli
end interface
contains
integer function muli(a,b)
integer::a,b
muli=a*b
end
end module m
"
);
c!(
    gen_if_08,
    "module m
interface g
procedure f
end interface
contains
subroutine f()
end subroutine f
end module m
"
);
c!(
    gen_if_09,
    "module m
interface write(formatted)
module procedure wf
end interface
contains
subroutine wf()
end subroutine wf
end module m
"
);
c!(
    gen_if_10,
    "module m
interface read(formatted)
module procedure rf
end interface
contains
subroutine rf()
end subroutine rf
end module m
"
);
c!(
    gen_if_11,
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
    gen_if_12,
    "module m
interface operator(//)
module procedure cat
end interface
contains
character(len=2) function cat(a,b)
character(len=1)::a,b
cat=a//b
end
end module m
"
);
c!(
    gen_if_13,
    "module m
interface assignment(=)
module procedure asgr
end interface
contains
subroutine asgr(a,b)
real::a
integer::b
a=real(b)
end
end module m
"
);
c!(
    gen_if_14,
    "module m
interface g
module procedure i1,i2
end interface
contains
integer function i1()
i1=1
end
integer function i2()
i2=2
end
end module m
"
);
c!(
    gen_if_15,
    "module m
interface g
module procedure s1
end interface
contains
subroutine s1(a,b)
integer::a,b
end
end module m
"
);
c!(
    gen_if_16,
    "module m
interface operator(.foo.)
module procedure foo
end interface
contains
logical function foo(a,b)
logical::a,b
foo=a.and.b
end
end module m
"
);
c!(
    gen_if_17,
    "module m
abstract interface
subroutine s(x)
integer::x
end
end interface
end module m
"
);
c!(
    gen_if_18,
    "module m
interface g
module procedure ss
end interface
contains
subroutine ss()
print *,1
end
end module m
"
);
c!(
    gen_if_19,
    "module m
interface g
module procedure fs
end interface
contains
integer function fs()
fs=1
end
end module m
"
);
c!(
    gen_if_20,
    "module m
interface operator(==)
module procedure eqi
end interface
contains
logical function eqi(a,b)
integer::a,b
eqi=a==b
end
end module m
"
);
