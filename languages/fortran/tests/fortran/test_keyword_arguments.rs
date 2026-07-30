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
    kw_01,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(x=1,y=2)
end program p
"
);
c!(
    kw_02,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(y=2,x=1)
end program p
"
);
c!(
    kw_03,
    "subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(z=3,x=1,y=2)
end program p
"
);
c!(
    kw_04,
    "subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(1,y=2,z=3)
end program p
"
);
c!(
    kw_05,
    "subroutine s(x,y)
integer, optional::x,y
end
program p
call s(y=2)
end program p
"
);
c!(
    kw_06,
    "subroutine s(x,y)
integer, optional::x,y
end
program p
call s(x=1)
end program p
"
);
c!(
    kw_07,
    "subroutine s(a,b,c)
real::a,b,c
end
program p
call s(c=3.0,b=2.0,a=1.0)
end program p
"
);
c!(
    kw_08,
    "subroutine s(i,j)
integer::i,j
end
program p
call s(j=2,i=1)
end program p
"
);
c!(
    kw_09,
    "subroutine s(x)
integer::x
end
program p
call s(x=1)
end program p
"
);
c!(
    kw_10,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(1,y=2)
end program p
"
);
c!(
    kw_11,
    "subroutine s(x,y,z,w)
integer::x,y,z,w
end
program p
call s(w=4,z=3,y=2,x=1)
end program p
"
);
c!(
    kw_12,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(x=1, y=2)
end program p
"
);
c!(
    kw_13,
    "subroutine s(flag,val)
logical::flag
integer::val
end
program p
call s(val=2, flag=.true.)
end program p
"
);
c!(
    kw_14,
    "subroutine s(x,y)
character(len=*)::x
automatic integer::y
end
program p
call s(y=1, x='a')
end program p
"
);
c!(
    kw_15,
    "subroutine s(a,b)
integer::a,b
end
program p
call s(b=2, a=1)
end program p
"
);
c!(
    kw_16,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(x=1, y=2)
end program p
"
);
c!(
    kw_17,
    "subroutine s(x,y)
integer, value::x,y
end
program p
call s(y=2,x=1)
end program p
"
);
c!(
    kw_18,
    "subroutine s(x,y)
integer, intent(in)::x,y
end
program p
call s(y=2,x=1)
end program p
"
);
c!(
    kw_19,
    "subroutine s(x,y)
integer, optional::x,y
end
program p
call s(x=1,y=2)
end program p
"
);
c!(
    kw_20,
    "subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(x=1,z=3,y=2)
end program p
"
);

c!(
    kw_21,
    "subroutine s(x,y,z,w)
integer::x, y, z, w
end
program p
call s(1, 2, w=4, z=3)
end program p
"
);

c!(
    kw_22,
    "subroutine s(name, value, scale)
character(len=*)::name
integer::value, scale
end
program p
call s(name='abc', value=3, scale=2)
end program p
"
);

c!(
    kw_23,
    "subroutine s(arr, n)
integer::arr(:)\ninteger::n
end
program p
call s(arr=[1,2,3], n=3)
end program p
"
);

c!(
    kw_24,
    "subroutine s(x,y,z)
integer, optional :: x, y, z
end
program p
call s(x=1, z=3)
end program p
"
);

c!(
    kw_25,
    "subroutine s(x, y)
real, intent(in) :: x
integer, intent(out) :: y
end
program p
integer :: y
call s(3.14, y)
end program p
"
);

c!(
    kw_26,
    "module m
interface
subroutine s(x, y)
integer::x,y
end subroutine
end interface
end module m
program p
use m
call s(y=2, x=1)
end program p
"
);
