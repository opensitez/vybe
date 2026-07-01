use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(nopass_arguments_01,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s()
end
end module m
");
c!(nopass_arguments_02,"module m
type::t
contains
procedure,nopass::f
end type
contains
integer function f()
f=1
end
end module m
");
c!(nopass_arguments_03,"module m
type::t
contains
procedure,nopass::s1
procedure,nopass::s2
end type
contains
subroutine s1()
end
subroutine s2()
end
end module m
");
c!(nopass_arguments_04,"module m
type::t
contains
procedure,nopass::set
end type
contains
subroutine set(x)
integer::x
end
end module m
");
c!(nopass_arguments_05,"module m
type::t
contains
procedure,nopass::mk
end type
contains
function mk() result(r)
integer :: r
r=1
end
end module m
");
c!(nopass_arguments_06,"module m
type::t
contains
procedure,nopass::show
end type
contains
subroutine show()
print *,1
end
end module m
");
c!(nopass_arguments_07,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(x)
real::x
end
end module m
");
c!(nopass_arguments_08,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(x,y)
integer::x,y
end
end module m
");
c!(nopass_arguments_09,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(c)
character(len=*)::c
end
end module m
");
c!(nopass_arguments_10,"module m
type::t
contains
procedure,nopass::s
end type
contains
subroutine s(l)
logical::l
end
end module m
");