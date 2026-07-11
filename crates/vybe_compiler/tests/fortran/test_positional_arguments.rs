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
    positional_arguments_01,
    "subroutine s(x)
integer::x
end
program p
call s(1)
end program p
"
);
c!(
    positional_arguments_02,
    "subroutine s(x,y)
integer::x,y
end
program p
call s(1,2)
end program p
"
);
c!(
    positional_arguments_03,
    "subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(1,2,3)
end program p
"
);
c!(
    positional_arguments_04,
    "subroutine s(x)
real::x
end
program p
call s(1.0)
end program p
"
);
c!(
    positional_arguments_05,
    "subroutine s(x)
character(len=*)::x
end
program p
call s('abc')
end program p
"
);
c!(
    positional_arguments_06,
    "subroutine s(x)
logical::x
end
program p
call s(.true.)
end program p
"
);
c!(
    positional_arguments_07,
    "subroutine s(x)
complex::x
end
program p
call s((1.0,2.0))
end program p
"
);
c!(
    positional_arguments_08,
    "subroutine s(a)
integer::a(2)
end
program p
integer::a(2)=[1,2]
call s(a)
end program p
"
);
c!(
    positional_arguments_09,
    "subroutine s(x,y)
integer::x
real::y
end
program p
call s(1,2.0)
end program p
"
);
c!(
    positional_arguments_10,
    "subroutine s(x,y)
character(len=*)::x
integer::y
end
program p
call s('a',1)
end program p
"
);
