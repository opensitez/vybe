use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(real_kinds_01,"program p
real(kind=4) :: x
print *, x
end program p
");
c!(real_kinds_02,"program p
real(kind=8) :: x
print *, x
end program p
");
c!(real_kinds_03,"program p
print *, selected_real_kind(6)
end program p
");
c!(real_kinds_04,"program p
print *, selected_real_kind(15)
end program p
");
c!(real_kinds_05,"program p
real(kind=8) :: x=1.0_8
print *, x
end program p
");
c!(real_kinds_06,"program p
real(kind=4) :: a=1.0_4,b=2.0_4
print *, a+b
end program p
");
c!(real_kinds_07,"program p
real(kind=8) :: a=4.0_8
print *, sqrt(a)
end program p
");
c!(real_kinds_08,"program p
real(kind=8) :: a=1.5_8
print *, floor(a)
end program p
");
c!(real_kinds_09,"program p
real(kind=8) :: a=1.5_8
print *, ceiling(a)
end program p
");
c!(real_kinds_10,"program p
real(kind=8) :: a=1.0_8,b=2.0_8
print *, nearest(a,b)
end program p
");