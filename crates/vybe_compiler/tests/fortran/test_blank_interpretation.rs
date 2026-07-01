use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(blank_interpretation_01,"program p
character(len=3) :: s='a b'
print *, s
end program p
");
c!(blank_interpretation_02,"program p
print *, scan('a b',' ')
end program p
");
c!(blank_interpretation_03,"program p
print *, verify('a b','ab')
end program p
");
c!(blank_interpretation_04,"program p
character(len=5) :: s='  abc'
print *, adjustl(s)
end program p
");
c!(blank_interpretation_05,"program p
character(len=5) :: s='abc  '
print *, adjustr(s)
end program p
");
c!(blank_interpretation_06,"program p
character(len=5) :: s='a   '
print *, len_trim(s)
end program p
");
c!(blank_interpretation_07,"program p
character(len=5) :: s='     '
print *, len_trim(s)
end program p
");
c!(blank_interpretation_08,"program p
character(len=6) :: s='ab cd '
print *, trim(s)
end program p
");
c!(blank_interpretation_09,"program p
character(len=6) :: s='ab cd '
print *, s(3:4)
end program p
");
c!(blank_interpretation_10,"program p
character(len=5) :: s=' a  '
print *, s
end program p
");