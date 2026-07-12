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
    character_assignment_01,
    "program p
character(len=4) :: s
s='ab'
print *, s
end program p
"
);
c!(
    character_assignment_02,
    "program p
character(len=4) :: s='abcd'
s='xy'
print *, s
end program p
"
);
c!(
    character_assignment_03,
    "program p
character(len=6) :: s
s='hello '
print *, s
end program p
"
);
c!(
    character_assignment_04,
    "program p
character(len=5) :: s='hello'
s(1:1)='H'
print *, s
end program p
"
);
c!(
    character_assignment_05,
    "program p
character(len=6) :: s='abcdef'
s(2:3)='ZZ'
print *, s
end program p
"
);
c!(
    character_assignment_06,
    "program p
character(len=:), allocatable :: s
allocate(character(len=3) :: s)
s='abc'
print *, s
end program p
"
);
c!(
    character_assignment_07,
    "program p
character(len=4) :: a(2)
a(1)='ab'
a(2)='cd'
print *, a
end program p
"
);
c!(
    character_assignment_08,
    "program p
character(len=8) :: s
s='ab'//'cd'
print *, s
end program p
"
);
c!(
    character_assignment_09,
    "program p
character(len=5) :: s='abcde'
s = adjustl('  xy')
print *, s
end program p
"
);
c!(
    character_assignment_10,
    "program p
character(len=5) :: s
s = repeat('x',3)
print *, s
end program p
"
);
