use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(generic_resolution_01,"module m
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
");
c!(generic_resolution_02,"module m
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
");
c!(generic_resolution_03,"module m
interface operator(+)
module procedure addi
end interface
contains
integer function addi(a,b)
integer::a,b
addi=a+b
end
end module m
");
c!(generic_resolution_04,"module m
interface assignment(=)
module procedure asg
end interface
contains
subroutine asg(a,b)
integer::a,b
a=b
end
end module m
");
c!(generic_resolution_05,"module m
interface g
module procedure si
end interface
contains
subroutine si(i)
integer::i
end
end module m
program p
use m
call g(1)
end program p
");
c!(generic_resolution_06,"module m
interface g
module procedure sr
end interface
contains
subroutine sr(r)
real::r
end
end module m
program p
use m
call g(1.0)
end program p
");
c!(generic_resolution_07,"module m
interface g
module procedure ss
end interface
contains
subroutine ss(s)
character(len=*)::s
end
end module m
");
c!(generic_resolution_08,"module m
interface operator(-)
module procedure subi
end interface
contains
integer function subi(a,b)
integer::a,b
subi=a-b
end
end module m
");
c!(generic_resolution_09,"module m
interface operator(*)
module procedure muli
end interface
contains
integer function muli(a,b)
integer::a,b
muli=a*b
end
end module m
");
c!(generic_resolution_10,"module m
interface g
module procedure li
end interface
contains
subroutine li(l)
logical::l
end
end module m
");