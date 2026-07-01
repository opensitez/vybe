use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(character_collation_01,"program p
logical :: l
l='a'<'b'
print *, l
end program p
");
c!(character_collation_02,"program p
logical :: l
l='A'<'B'
print *, l
end program p
");
c!(character_collation_03,"program p
logical :: l
l='A'<'a'
print *, l
end program p
");
c!(character_collation_04,"program p
logical :: l
l='0'<'9'
print *, l
end program p
");
c!(character_collation_05,"program p
logical :: l
l='abc'<'abd'
print *, l
end program p
");
c!(character_collation_06,"program p
logical :: l
l='abc'<='abc'
print *, l
end program p
");
c!(character_collation_07,"program p
logical :: l
l='xyz'>='xy'
print *, l
end program p
");
c!(character_collation_08,"program p
logical :: l
l='b'>'a'
print *, l
end program p
");
c!(character_collation_09,"program p
logical :: l
l='a'/='b'
print *, l
end program p
");
c!(character_collation_10,"program p
logical :: l
l='same'=='same'
print *, l
end program p
");