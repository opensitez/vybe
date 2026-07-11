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
    pass_arguments_01,
    "module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t)::this
end
end module m
"
);
c!(
    pass_arguments_02,
    "module m
type::t
contains
procedure,pass(self)::s
end type
contains
subroutine s(self)
class(t)::self
end
end module m
"
);
c!(
    pass_arguments_03,
    "module m
type::t
contains
procedure,pass(arg)::s
end type
contains
subroutine s(arg)
class(t)::arg
end
end module m
"
);
c!(
    pass_arguments_04,
    "module m
type::t
contains
procedure,pass::s1
procedure,pass::s2
end type
contains
subroutine s1(this)
class(t)::this
end
subroutine s2(this)
class(t)::this
end
end module m
"
);
c!(
    pass_arguments_05,
    "module m
type::t
contains
procedure,pass::show
end type
contains
subroutine show(this)
class(t)::this
print *,1
end
end module m
"
);
c!(
    pass_arguments_06,
    "module m
type::t
contains
procedure,pass::get
end type
contains
integer function get(this)
class(t)::this
get=1
end
end module m
"
);
c!(
    pass_arguments_07,
    "module m
type::t
contains
procedure,pass::set
end type
contains
subroutine set(this,x)
class(t)::this
integer::x
end
end module m
"
);
c!(
    pass_arguments_08,
    "module m
type::t
contains
procedure,pass::a
procedure,pass::b
end type
contains
subroutine a(this)
class(t)::this
end
subroutine b(this)
class(t)::this
end
end module m
"
);
c!(
    pass_arguments_09,
    "module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t), intent(inout) :: this
end
end module m
"
);
c!(
    pass_arguments_10,
    "module m
type::t
contains
procedure,pass::s
end type
contains
subroutine s(this)
class(t), intent(in) :: this
end
end module m
"
);
