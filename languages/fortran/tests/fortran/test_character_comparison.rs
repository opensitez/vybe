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
    character_comparison_01,
    "program p
logical :: l
l='a'=='a'
print *, l
end program p
"
);
c!(
    character_comparison_02,
    "program p
logical :: l
l='a'/='b'
print *, l
end program p
"
);
c!(
    character_comparison_03,
    "program p
logical :: l
l='a'<'b'
print *, l
end program p
"
);
c!(
    character_comparison_04,
    "program p
logical :: l
l='b'>'a'
print *, l
end program p
"
);
c!(
    character_comparison_05,
    "program p
logical :: l
l='a'<='b'
print *, l
end program p
"
);
c!(
    character_comparison_06,
    "program p
logical :: l
l='b'>='a'
print *, l
end program p
"
);
c!(
    character_comparison_07,
    "program p
logical :: l
l='abc'=='abc'
print *, l
end program p
"
);
c!(
    character_comparison_08,
    "program p
logical :: l
l='abc'/='abd'
print *, l
end program p
"
);
c!(
    character_comparison_09,
    "program p
logical :: l
l='A'/='a'
print *, l
end program p
"
);
c!(
    character_comparison_10,
    "program p
character(len=3) :: a='abc', b='abd'
print *, a < b
end program p
"
);
