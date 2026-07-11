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
    padding_rules_01,
    "program p
character(len=5) :: s
s='a'
print *, s
end program p
"
);
c!(
    padding_rules_02,
    "program p
character(len=5) :: s='ab'
print *, s
end program p
"
);
c!(
    padding_rules_03,
    "program p
character(len=3) :: s='abcdef'
print *, s
end program p
"
);
c!(
    padding_rules_04,
    "program p
character(len=4) :: s
s='xy'//'z'
print *, s
end program p
"
);
c!(
    padding_rules_05,
    "program p
character(len=6) :: s='hi'
print *, len_trim(s)
end program p
"
);
c!(
    padding_rules_06,
    "program p
character(len=4) :: a(2)
a(1)='a'
a(2)='bc'
print *, a
end program p
"
);
c!(
    padding_rules_07,
    "program p
character(len=5) :: s='abc'
print *, trim(s)
end program p
"
);
c!(
    padding_rules_08,
    "program p
character(len=5) :: s='abc'
print *, adjustl(s)
end program p
"
);
c!(
    padding_rules_09,
    "program p
character(len=5) :: s='abc'
print *, adjustr(s)
end program p
"
);
c!(
    padding_rules_10,
    "program p
character(len=8) :: s
s = repeat('x',3)
print *, s
end program p
"
);
