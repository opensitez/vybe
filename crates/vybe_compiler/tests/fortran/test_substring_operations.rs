use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(substring_operations_01,"program p
character(len=5) :: s='hello'
print *, s(1:2)
end program p
");
c!(substring_operations_02,"program p
character(len=5) :: s='hello'
print *, s(2:4)
end program p
");
c!(substring_operations_03,"program p
character(len=5) :: s='hello'
print *, s(:3)
end program p
");
c!(substring_operations_04,"program p
character(len=5) :: s='hello'
print *, s(3:)
end program p
");
c!(substring_operations_05,"program p
character(len=6) :: s='abcdef'
s(2:3)='ZZ'
print *, s
end program p
");
c!(substring_operations_06,"program p
character(len=6) :: s='abcdef'
s(1:1)='X'
print *, s
end program p
");
c!(substring_operations_07,"program p
character(len=6) :: s='abcdef'
print *, s(6:6)
end program p
");
c!(substring_operations_08,"program p
character(len=6) :: s='abcdef'
print *, len(s(2:5))
end program p
");
c!(substring_operations_09,"program p
character(len=6) :: s='abcdef'
print *, index(s(2:), 'de')
end program p
");
c!(substring_operations_10,"program p
character(len=6) :: s='abcdef'
print *, scan(s(2:5), 'd')
end program p
");