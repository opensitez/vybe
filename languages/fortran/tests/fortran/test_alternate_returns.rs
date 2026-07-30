use super::helpers::{compile_ok, parse_ok, run_prints};
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    alternate_returns_01,
    "subroutine s(*,*)
return 1
end
"
);
c!(
    alternate_returns_02,
    "subroutine s(*,*)
return 2
end
"
);
c!(
    alternate_returns_03,
    "subroutine s(*,*)
return
end
"
);
c!(
    alternate_returns_04,
    "program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return 1
end
"
);
c!(
    alternate_returns_05,
    "program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return 2
end
"
);
c!(
    alternate_returns_06,
    "program p
call s(*10,*20)
10 continue
20 continue
end program p
subroutine s(*,*)
return
end
"
);
c!(
    alternate_returns_07,
    "subroutine s(x,*,*)
integer::x
return 1
end
"
);
c!(
    alternate_returns_08,
    "subroutine s(x,*,*)
integer::x
return 2
end
"
);
c!(
    alternate_returns_09,
    "program p
integer::x=1
call s(x,*10,*20)
10 continue
20 continue
end program p
subroutine s(x,*,*)
integer::x
return 1
end
"
);
c!(
    alternate_returns_10,
    "program p
integer::x=2
call s(x,*10,*20)
10 continue
20 continue
end program p
subroutine s(x,*,*)
integer::x
return 2
end
"
);

#[test]
fn alternate_returns_runtime_first_alternate_is_taken_on_return_1() {
    assert_eq!(
        run_prints(
            "program p\n\
integer :: x\n\
x = 0\n\
call s(*10,*20)\n\
print *, x\n\
10 x = 2\n\
print *, x\n\
20 continue\n\
end program p\n\
subroutine s(*,*)\n\
return 1\n\
end"
        ),
        vec!["2"]
    );
}

#[test]
fn alternate_returns_runtime_second_alternate_is_taken_on_return_2() {
    assert_eq!(
        run_prints(
            "program p\n\
integer :: x\n\
x = 0\n\
call s(*10,*20)\n\
print *, x\n\
10 x = 2\n\
print *, x\n\
20 x = 3\n\
print *, x\n\
end program p\n\
subroutine s(*,*)\n\
return 2\n\
end"
        ),
        vec!["3"]
    );
}

#[test]
fn alternate_returns_runtime_plain_return_defaults_to_call_fallthrough() {
    assert_eq!(
        run_prints(
            "program p\n\
integer :: x\n\
x = 0\n\
call s(*10,*20)\n\
x = 1\n\
print *, x\n\
10 print *, x\n\
20 continue\n\
end program p\n\
subroutine s(*,*)\n\
return\n\
end"
        ),
        vec!["1"]
    );
}

#[test]
fn alternate_returns_syntax_rejects_malformed_formal_list() {
    assert!(!parse_ok("subroutine s(*,\nend"));
}
